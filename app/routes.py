import csv
import copy
import hashlib
import io
import json
import math
import os
import re
import secrets
import smtplib
import ssl
import stat
import sqlite3
import shutil
import subprocess
import tempfile
import time
import uuid
import urllib.error
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from functools import wraps
from html import escape
from pathlib import Path
from urllib.parse import urlsplit
from flask import Blueprint, render_template, redirect, url_for, request, flash, abort, current_app, session, send_from_directory, send_file, Response, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from itsdangerous import BadSignature, URLSafeTimedSerializer
from sqlalchemy import and_, literal, or_, select, union_all
from sqlalchemy.exc import IntegrityError, SQLAlchemyError
from sqlalchemy.schema import CreateTable
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from openpyxl.utils.exceptions import InvalidFileException
import pymupdf
import qrcode
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException
from .extensions import csrf, db
from .onlyoffice import InvalidOnlyOfficeToken, decode_onlyoffice_token, encode_onlyoffice_token
from .models import User, Team, TeamMember, Domain, Evaluation, AdminEvaluation, MentorEvaluation, TeamTask, WorkTask, WorkTaskAssignment, TaskDeadlineReminder, TeamChatMessage, AdminChatMessage, Nomination, ProjectVideo, TeamSubmission, Winner, DomainResultPublication, UserNotification, CertificateType, CertificateTemplate, Certificate, Announcement, AuditLog, JudgeAssignment, MentorTeamAssignment, CoordinatorDomainAssignment, CoordinatorTeamAssignment, LoginAttempt, SiteSetting, Attendance
from .rich_text import sanitize_rich_text

main = Blueprint("main", __name__)

USER_ROLES = {"admin", "team_leader", "team_member", "mentor", "coordinator", "judge"}
REGISTRATION_STATUSES = {"Pending", "Approved", "Rejected", "Changes Required"}
LOGIN_WINDOW = timedelta(minutes=15)
MAX_LOGIN_ATTEMPTS = 10
CHAT_EDIT_WINDOW = timedelta(minutes=5)
CHAT_FILE_MAX_BYTES = 25 * 1024 * 1024
TASK_FILE_MAX_BYTES = 25 * 1024 * 1024
TEAM_SUBMISSION_MAX_BYTES = 25 * 1024 * 1024
TEAM_SUBMISSION_EXTENSIONS = {
    "ppt": {".ppt", ".pptx"},
    "documentation": {".pdf", ".doc", ".docx"},
}
TASK_CREATOR_ROLES = {"admin", "team_leader", "mentor", "coordinator"}
TASK_ASSIGNABLE_ROLES = {
    "admin": {"admin", "mentor", "coordinator", "team_leader", "team_member"},
    "mentor": {"coordinator", "team_leader", "team_member"},
    "coordinator": {"team_leader", "team_member"},
    "team_leader": {"team_member"},
}
TASK_STATUSES = {"Open", "In Progress", "Submitted", "Completed"}


def roles_required(*roles):
    def deco(fn):
        @wraps(fn)
        @login_required
        def wrapper(*args, **kwargs):
            if current_user.role not in roles:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return deco

@main.before_app_request
def enforce_admin_impersonation():
    admin_id = session.get("admin_impersonation_admin_id")
    if not admin_id:
        return None
    if session.get("admin_impersonation_mode") == "view_only" and (
        request.method not in {"GET", "HEAD", "OPTIONS"}
        and request.endpoint not in {
            "main.end_admin_impersonation",
            "main.logout",
            "main.upgrade_admin_impersonation",
        }
    ):
        abort(403, description="This admin impersonation session is view-only.")
    return None


@main.after_app_request
def audit_admin_impersonation_request(response):
    admin_id = session.get("admin_impersonation_admin_id")
    if (
        admin_id
        and request.method in {"POST", "PUT", "PATCH", "DELETE"}
        and request.endpoint not in {
            "main.end_admin_impersonation",
            "main.logout",
            "main.upgrade_admin_impersonation",
        }
        and current_user.is_authenticated
    ):
        db.session.add(AuditLog(
            user_id=admin_id,
            action=f"Impersonated {request.method} request"[:200],
            details=(
                f"Acting as user ID {current_user.id} "
                f"({current_user.username}); endpoint={request.endpoint or 'unknown'}; "
                f"status={response.status_code}"
            ),
        ))
        db.session.commit()
    return response


@main.before_app_request
def process_task_deadline_reminders():
    if request.endpoint == "static":
        return None
    deliver_due_task_reminders()
    return None


def log(action, details=""):
    admin_id = session.get("admin_impersonation_admin_id")
    if admin_id:
        details = (
            f"Acting as user ID {current_user.id} ({current_user.username}). "
            f"{details}"
        ).strip()
        db.session.add(AuditLog(user_id=admin_id, action=action, details=details))
    else:
        db.session.add(AuditLog(user_id=current_user.id, action=action, details=details))


def parse_scores(limits):
    scores = {}
    for field, maximum in limits.items():
        try:
            score = float(request.form[field])
        except (KeyError, TypeError, ValueError):
            abort(400, description=f"Invalid score for {field}.")
        if (
            not math.isfinite(score)
            or score < 0
            or score > maximum
            or not (score * 2).is_integer()
        ):
            abort(400, description=f"Score for {field} must be between 0 and {maximum}.")
        scores[field] = score
    return scores


def assigned_teams_for_user(user):
    if user.role == "mentor":
        return [
            assignment.team
            for assignment in MentorTeamAssignment.query.filter_by(mentor_id=user.id)
            .join(Team, MentorTeamAssignment.team_id == Team.id)
            .order_by(Team.team_code)
        ]
    if user.role == "coordinator":
        assigned = [
            assignment.team
            for assignment in CoordinatorTeamAssignment.query.filter_by(
                coordinator_id=user.id
            )
            .join(Team, CoordinatorTeamAssignment.team_id == Team.id)
            .order_by(Team.team_code)
        ]
        domain_ids = [
            assignment.domain_id
            for assignment in CoordinatorDomainAssignment.query.filter_by(
                coordinator_id=user.id
            )
        ]
        if domain_ids:
            assigned.extend(
                Team.query.filter(Team.domain_id.in_(domain_ids))
                .order_by(Team.team_code).all()
            )
        return list({team.id: team for team in assigned}.values())
    return []


def can_view_team(user, team):
    if user.role == "admin":
        return True
    if user.role in {"team_leader", "team_member"}:
        return user.team_id == team.id
    if user.role == "mentor":
        return MentorTeamAssignment.query.filter_by(
            mentor_id=user.id, team_id=team.id
        ).first() is not None
    if user.role == "coordinator":
        if CoordinatorTeamAssignment.query.filter_by(
            coordinator_id=user.id, team_id=team.id
        ).first() is not None:
            return True
        return CoordinatorDomainAssignment.query.filter_by(
            coordinator_id=user.id, domain_id=team.domain_id
        ).first() is not None
    return False


def can_use_team_workspace(user, team):
    return user.role in {
        "admin", "team_leader", "team_member", "mentor", "coordinator"
    } and can_view_team(user, team)


def collaboration_users(team):
    users = User.query.filter(
        User.active.is_(True),
        User.role.in_(("team_leader", "team_member", "mentor", "coordinator")),
    ).order_by(User.display_name, User.username).all()
    return [user for user in users if can_view_team(user, team)]


def task_assignees_for_role(role, users):
    allowed_roles = TASK_ASSIGNABLE_ROLES.get(role, set())
    return [user for user in users if user.role in allowed_roles]


def save_task_upload(upload):
    if not upload or not upload.filename:
        return None, None
    original_name = os.path.basename(upload.filename.replace("\\", "/")).strip()
    original_name = "".join(
        character for character in original_name if character.isprintable()
    )[:255]
    if not original_name:
        abort(400, description="Choose a file with a valid name.")
    contents = upload.stream.read(TASK_FILE_MAX_BYTES + 1)
    if len(contents) > TASK_FILE_MAX_BYTES:
        abort(413, description="Task files cannot exceed 25 MB.")
    storage_name = uuid.uuid4().hex
    upload_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "task-files")
    os.makedirs(upload_directory, exist_ok=True)
    with open(os.path.join(upload_directory, storage_name), "wb") as file_stream:
        file_stream.write(contents)
    return storage_name, original_name


def user_notification_email(user):
    account_email = (user.email or "").strip()
    if account_email:
        return account_email
    member = user.member_record
    return (member.email or "").strip() if member else ""


def user_notification_phone(user):
    account_phone = (user.phone or "").strip()
    if account_phone:
        return account_phone
    member = user.member_record
    return (member.phone or "").strip() if member else ""


def task_notification_configuration():
    smtp_host = current_app.config.get("TASK_EMAIL_SMTP_HOST", "").strip()
    smtp_username = current_app.config.get("TASK_EMAIL_SMTP_USERNAME", "").strip()
    smtp_password = current_app.config.get("TASK_EMAIL_SMTP_PASSWORD", "").strip()
    sender = current_app.config.get("TASK_EMAIL_FROM", "").strip()
    variable_name = current_app.config.get(
        "TASK_SMS_VARIABLE_NAME", "VAR1"
    ).strip()
    sms_ready = bool(
        current_app.config.get("TASK_SMS_AUTH_KEY", "").strip()
        and current_app.config.get("TASK_SMS_TEMPLATE_ID", "").strip()
        and re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,30}", variable_name)
    )
    active_users = User.query.filter_by(active=True).all()
    return {
        "email_ready": bool(
            smtp_host
            and sender
            and bool(smtp_username) == bool(smtp_password)
        ),
        "sms_ready": sms_ready,
        "users_missing_email": sum(
            1 for user in active_users if not user_notification_email(user)
        ),
        "users_missing_phone": sum(
            1 for user in active_users
            if not normalize_task_sms_phone(user_notification_phone(user))
        ),
    }


def send_task_assignment_emails(task, recipients, reminder_label=None):
    recipients_with_email = [
        (user, user_notification_email(user))
        for user in recipients
        if user_notification_email(user)
    ]
    missing_email = len(recipients) - len(recipients_with_email)
    host = current_app.config.get("TASK_EMAIL_SMTP_HOST", "").strip()
    if not recipients_with_email:
        return {"sent": 0, "failed": 0, "not_configured": 0, "missing_email": missing_email}
    if not host:
        return {
            "sent": 0,
            "failed": 0,
            "not_configured": len(recipients_with_email),
            "missing_email": missing_email,
        }

    port = current_app.config["TASK_EMAIL_SMTP_PORT"]
    username = current_app.config.get("TASK_EMAIL_SMTP_USERNAME", "")
    password = current_app.config.get("TASK_EMAIL_SMTP_PASSWORD", "")
    sender = current_app.config["TASK_EMAIL_FROM"]
    result = {"sent": 0, "failed": 0, "not_configured": 0, "missing_email": missing_email}
    try:
        smtp_class = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
        with smtp_class(host, port, timeout=15) as server:
            if port != 465:
                server.starttls(context=ssl.create_default_context())
            if username or password:
                if not username or not password:
                    raise smtplib.SMTPException("SMTP username and password must both be configured.")
                server.login(username, password)
            for recipient, recipient_email in recipients_with_email:
                message = EmailMessage()
                message["Subject"] = (
                    f"Task deadline {reminder_label}: {task.title}"
                    if reminder_label
                    else f"Task assigned: {task.title}"
                )
                message["From"] = sender
                message["To"] = recipient_email
                task_scope = (
                    f"Project {task.team.team_code}"
                    if task.team else "Club-wide"
                )
                task_page_url = url_for(
                    "main.team_workspace",
                    team_id=task.team_id,
                    _external=True,
                ) if isinstance(task, TeamTask) else url_for(
                    "main.work_tasks",
                    _external=True,
                )
                message.set_content(
                    f"Hello {recipient.display_name or recipient.username},\n\n" + (
                        f"Reminder: this task is due {reminder_label}.\n\n"
                        if reminder_label
                        else "A new task has been assigned to you.\n\n"
                    )
                    + f"Task: {task.title}\n"
                    f"Assigned by: {task.assigner.display_name or task.assigner.username}\n"
                    f"Scope: {task_scope}\n"
                    f"Instructions: {task.description or 'No additional instructions provided.'}\n"
                    f"Due date: {task.due_date.strftime('%d %b %Y') if task.due_date else 'Not specified'}\n\n"
                    "Sign in to FUTURE-TECH INNOVISION 2026 to update your task status.\n"
                    f"Task workspace: {task_page_url}\n\n"
                    f"Future-Tech Club contact: {current_app.config['TASK_CONTACT_PHONE']} | "
                    f"{sender}"
                )
                try:
                    server.send_message(message)
                    result["sent"] += 1
                except (smtplib.SMTPException, OSError) as exc:
                    result["failed"] += 1
                    current_app.logger.error(
                        "Could not email task assignment to account %s: %s",
                        recipient.id,
                        exc,
                    )
    except (smtplib.SMTPException, OSError, ssl.SSLError) as exc:
        result["failed"] = len(recipients_with_email) - result["sent"]
        current_app.logger.error("Task email delivery failed: %s", exc)
    return result


def send_certificate_mail_merge_emails(template, recipients):
    host = current_app.config.get("TASK_EMAIL_SMTP_HOST", "").strip()
    port = current_app.config["TASK_EMAIL_SMTP_PORT"]
    username = current_app.config.get("TASK_EMAIL_SMTP_USERNAME", "").strip()
    password = current_app.config.get("TASK_EMAIL_SMTP_PASSWORD", "").strip()
    sender = current_app.config.get("TASK_EMAIL_FROM", "").strip()
    result = {"sent": 0, "failed": 0}
    try:
        smtp_class = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
        with smtp_class(host, port, timeout=15) as server:
            if port != 465:
                server.starttls(context=ssl.create_default_context())
            if username or password:
                server.login(username, password)
            for user, recipient_email, certificate in recipients:
                try:
                    if certificate.file_path:
                        if os.path.basename(certificate.file_path) != certificate.file_path:
                            raise ValueError("The uploaded certificate file path is invalid.")
                        certificate_path = os.path.join(
                            current_app.config["UPLOAD_FOLDER"],
                            "certificates",
                            certificate.file_path,
                        )
                        with open(certificate_path, "rb") as certificate_file:
                            pdf_data = certificate_file.read()
                    else:
                        pdf_data = render_template_certificate(certificate).getvalue()
                    try:
                        fields = json.loads(certificate.fields_json or "{}")
                    except (TypeError, json.JSONDecodeError):
                        fields = {}
                    if not isinstance(fields, dict):
                        fields = {}
                    message = EmailMessage()
                    message["Subject"] = (
                        f"{certificate.type.name} · "
                        f"{template.event_name if template else fields.get('event_name', 'FUTURE-TECH INNOVISION 2026')}"
                    )
                    message["From"] = sender
                    message["To"] = recipient_email
                    message.set_content(
                        f"Hello {certificate.recipient_name},\n\n"
                        f"Your {certificate.type.name} for "
                        f"{template.event_name if template else 'FUTURE-TECH INNOVISION 2026'} "
                        f"has been approved and dispatched. Certificate number: "
                        f"{certificate.certificate_id}. Your PDF is attached.\n\n"
                        "Regards,\nFuture-Tech Club"
                    )
                    message.add_attachment(
                        pdf_data,
                        maintype="application",
                        subtype="pdf",
                        filename=f"{certificate.certificate_id}.pdf",
                    )
                    server.send_message(message)
                    result["sent"] += 1
                except (
                    HTTPException,
                    OSError,
                    RuntimeError,
                    ValueError,
                    smtplib.SMTPException,
                ) as error:
                    result["failed"] += 1
                    current_app.logger.error(
                        "Could not mail certificate to account %s: %s",
                        user.id,
                        error,
                    )
    except (OSError, ssl.SSLError, smtplib.SMTPException) as error:
        result["failed"] = len(recipients) - result["sent"]
        current_app.logger.error("Certificate mail merge delivery failed: %s", error)
    return result


def normalize_task_sms_phone(phone):
    digits = re.sub(r"\D", "", phone or "")
    country_code = current_app.config.get("TASK_SMS_COUNTRY_CODE", "91")
    if not country_code.isdigit() or not country_code or country_code.startswith("0"):
        return None
    if digits.startswith("00"):
        digits = digits[2:]
    if len(digits) == 10:
        digits = country_code + digits
    elif len(digits) == 11 and digits.startswith("0"):
        digits = country_code + digits[1:]
    if not 8 <= len(digits) <= 15 or digits.startswith("0"):
        return None
    return digits


def send_task_assignment_sms(task, recipients):
    auth_key = current_app.config.get("TASK_SMS_AUTH_KEY", "").strip()
    template_id = current_app.config.get("TASK_SMS_TEMPLATE_ID", "").strip()
    variable_name = current_app.config.get("TASK_SMS_VARIABLE_NAME", "VAR1").strip()
    configured = bool(
        auth_key
        and template_id
        and re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,30}", variable_name)
    )
    if not configured:
        return {"accepted": 0, "failed": 0, "not_configured": len(recipients), "missing_phone": 0}

    sms_recipients = []
    missing_phone = 0
    for recipient in recipients:
        mobile = normalize_task_sms_phone(user_notification_phone(recipient))
        if not mobile:
            missing_phone += 1
            continue
        sms_recipients.append({
            "mobiles": mobile,
            variable_name: task.title[:80],
        })
    if not sms_recipients:
        return {"accepted": 0, "failed": 0, "not_configured": 0, "missing_phone": missing_phone}

    payload = json.dumps({
        "template_id": template_id,
        "short_url": "0",
        "realTimeResponse": "1",
        "recipients": sms_recipients,
    }).encode("utf-8")
    sms_request = urllib.request.Request(
        "https://control.msg91.com/api/v5/flow",
        data=payload,
        headers={
            "accept": "application/json",
            "authkey": auth_key,
            "content-type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(sms_request, timeout=15) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        current_app.logger.error("MSG91 task SMS request failed with HTTP %s.", exc.code)
        return {"accepted": 0, "failed": len(sms_recipients), "not_configured": 0, "missing_phone": missing_phone}
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, UnicodeDecodeError) as exc:
        current_app.logger.error("MSG91 task SMS request could not be completed: %s", exc)
        return {"accepted": 0, "failed": len(sms_recipients), "not_configured": 0, "missing_phone": missing_phone}

    response_data = result.get("data") if isinstance(result, dict) else None
    response_type = (
        result.get("type") if isinstance(result, dict) else None
    ) or (
        response_data.get("type") if isinstance(response_data, dict) else None
    )
    if response_type != "success":
        current_app.logger.error("MSG91 rejected the task SMS request.")
        return {"accepted": 0, "failed": len(sms_recipients), "not_configured": 0, "missing_phone": missing_phone}
    return {"accepted": len(sms_recipients), "failed": 0, "not_configured": 0, "missing_phone": missing_phone}


def task_for_user_or_abort(task_id):
    task = db.session.get(WorkTask, task_id)
    if not task:
        abort(404)
    if current_user.role == "admin" or task.assigner_id == current_user.id:
        return task
    if WorkTaskAssignment.query.filter_by(task_id=task.id, user_id=current_user.id).first():
        return task
    abort(404)


def save_chat_attachment():
    upload = request.files.get("attachment")
    if not upload or not upload.filename:
        return None, None

    original_name = os.path.basename(upload.filename.replace("\\", "/")).strip()
    original_name = "".join(character for character in original_name if character.isprintable())
    if not original_name:
        abort(400, description="Choose a file with a valid name.")
    original_name = original_name[:255]
    contents = upload.stream.read(CHAT_FILE_MAX_BYTES + 1)
    if len(contents) > CHAT_FILE_MAX_BYTES:
        abort(413, description="Chat attachments cannot exceed 25 MB.")

    storage_name = uuid.uuid4().hex
    upload_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "chat-files")
    os.makedirs(upload_directory, exist_ok=True)
    with open(os.path.join(upload_directory, storage_name), "wb") as attachment_file:
        attachment_file.write(contents)
    return storage_name, original_name


def serialize_chat_message(message, is_own):
    edit_remaining = max(
        0,
        int((CHAT_EDIT_WINDOW - (datetime.utcnow() - message.sent_at)).total_seconds()),
    )
    is_team_message = isinstance(message, TeamChatMessage)
    return {
        "id": message.id,
        "sender_id": message.sender_id,
        "sender": message.sender.display_name or message.sender.username,
        "sent_at": message.sent_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "body": message.body,
        "edited": message.edited_at is not None,
        "can_edit": is_own and edit_remaining > 0,
        "edit_remaining": edit_remaining if is_own else 0,
        "attachment_name": message.attachment_original_name,
        "attachment_url": (
            url_for("main.download_team_chat_attachment", message_id=message.id)
            if is_team_message and message.attachment_storage_name
            else url_for("main.download_admin_chat_attachment", message_id=message.id)
            if message.attachment_storage_name
            else None
        ),
        "edit_url": (
            url_for("main.edit_team_chat_message", message_id=message.id)
            if is_team_message else url_for(
                "main.edit_admin_chat_message"
                if current_user.role == "admin"
                else "main.edit_user_admin_chat_message",
                message_id=message.id,
            )
        ),
    }


def active_admin_account(user_id):
    user = db.session.get(User, user_id)
    if not user or user.role != "admin" or not user.active:
        abort(404)
    return user


def active_admin_chat_user(user_id):
    user = db.session.get(User, user_id)
    if not user or not user.active or user.id == current_user.id:
        abort(404)
    return user


def chat_contacts_for_user(user):
    if user.role == "admin":
        return User.query.filter(
            User.active.is_(True),
            User.id != user.id,
        ).order_by(User.role, User.display_name, User.username).all()

    teams = []
    global_roles = {"admin"}
    team_roles = set()
    if user.role == "mentor":
        global_roles.update({"mentor", "coordinator", "team_leader"})
        team_roles.add("team_member")
        teams = assigned_teams_for_user(user)
    elif user.role == "coordinator":
        global_roles.update({"mentor", "coordinator"})
        team_roles.update({"team_leader", "team_member"})
        teams = assigned_teams_for_user(user)
    elif user.role == "team_leader":
        global_roles.add("mentor")
        team_roles.update({"coordinator", "team_member"})
        team = db.session.get(Team, user.team_id) if user.team_id else None
        teams = [team] if team else []
    elif user.role == "team_member":
        global_roles.add("mentor")
        team_roles.update({"coordinator", "team_leader"})
        team = db.session.get(Team, user.team_id) if user.team_id else None
        teams = [team] if team else []

    contacts = User.query.filter(
        User.active.is_(True),
        User.id != user.id,
        User.role.in_(global_roles),
    ).all()
    if team_roles and teams:
        candidates = User.query.filter(
            User.active.is_(True),
            User.id != user.id,
            User.role.in_(team_roles),
        ).all()
        contacts.extend(
            candidate
            for candidate in candidates
            if any(can_view_team(candidate, team) for team in teams)
        )
    return sorted(
        {contact.id: contact for contact in contacts}.values(),
        key=lambda contact: (
            contact.role,
            (contact.display_name or "").casefold(),
            contact.username.casefold(),
        ),
    )


def active_chat_contact(user_id):
    user = db.session.get(User, user_id)
    if (
        not user
        or not user.active
        or user.id == current_user.id
        or user not in chat_contacts_for_user(current_user)
    ):
        abort(404)
    return user


def chat_upload_response(message):
    return {
        "message": serialize_chat_message(message, is_own=True),
    }


def approved_team_score_details(team):
    judges = Evaluation.query.filter_by(
        team_id=team.id, finalized=True, approved=True
    ).order_by(Evaluation.judge_id).all()
    admin_score = AdminEvaluation.query.filter_by(
        team_id=team.id, finalized=True
    ).first()
    mentor_scores = MentorEvaluation.query.filter_by(
        team_id=team.id, status="Approved"
    ).order_by(MentorEvaluation.submitted_at.desc()).all()
    return {
        "judges": judges,
        "admin_score": admin_score,
        "mentor_scores": mentor_scores,
        "total": (
            sum(score.total for score in judges) / len(judges) + admin_score.total
            if judges and admin_score else None
        ),
    }


def save_site_setting(key, value):
    setting = db.session.get(SiteSetting, key)
    if setting is None:
        setting = SiteSetting(key=key, value=value)
        db.session.add(setting)
    else:
        setting.value = value


def validate_image_upload(file):
    extensions = {
        ".png": b"\x89PNG\r\n\x1a\n",
        ".jpg": b"\xff\xd8\xff",
        ".jpeg": b"\xff\xd8\xff",
        ".webp": b"RIFF",
    }
    extension = os.path.splitext(file.filename or "")[1].lower()
    signature = extensions.get(extension)
    if not signature:
        return False
    file.stream.seek(0, os.SEEK_END)
    if file.stream.tell() > 5 * 1024 * 1024:
        file.stream.seek(0)
        return False
    file.stream.seek(8)
    webp_marker = file.stream.read(4) if extension == ".webp" else b""
    file.stream.seek(0)
    header = file.stream.read(len(signature))
    file.stream.seek(0)
    return header == signature and (
        extension != ".webp" or webp_marker == b"WEBP"
    )


def save_team_video(team, file):
    if not file or not file.filename:
        flash("Choose an MP4 video to upload.", "danger")
        return False
    if not file.filename.lower().endswith(".mp4"):
        flash("Please upload an MP4 video.", "danger")
        return False
    signature = file.stream.read(12)
    file.stream.seek(0)
    if len(signature) < 12 or signature[4:8] != b"ftyp":
        flash("The selected file does not appear to be a valid MP4 video.", "danger")
        return False

    filename = f"{team.team_code}_{uuid.uuid4().hex}.mp4"
    video_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "videos")
    os.makedirs(video_directory, exist_ok=True)
    path = os.path.join(video_directory, filename)
    file.save(path)
    existing_videos = ProjectVideo.query.filter_by(team_id=team.id).all()
    previous_filenames = [video.filename for video in existing_videos]
    video = existing_videos[0] if existing_videos else None
    if video:
        video.filename = filename
        video.uploaded_at = datetime.utcnow()
        for duplicate in existing_videos[1:]:
            db.session.delete(duplicate)
    else:
        db.session.add(ProjectVideo(team_id=team.id, filename=filename))
    return filename, previous_filenames


def remove_replaced_video_files(previous_filenames, current_filename):
    video_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "videos")
    for previous_filename in previous_filenames:
        if previous_filename == current_filename:
            continue
        if os.path.basename(previous_filename) != previous_filename:
            continue
        previous_path = os.path.join(video_directory, previous_filename)
        if os.path.isfile(previous_path):
            os.remove(previous_path)


def save_team_submission(team, submission_type, upload):
    allowed_extensions = TEAM_SUBMISSION_EXTENSIONS[submission_type]
    original_name = os.path.basename((upload.filename or "").replace("\\", "/")).strip()
    extension = os.path.splitext(original_name)[1].lower()
    if not original_name or extension not in allowed_extensions:
        allowed = ", ".join(sorted(allowed_extensions))
        flash(f"Choose a {submission_type} file with one of these extensions: {allowed}.", "danger")
        return None
    original_name = "".join(
        character for character in original_name if character.isprintable()
    )[:255]
    if not original_name:
        flash("Choose a file with a valid name.", "danger")
        return None
    contents = upload.stream.read(TEAM_SUBMISSION_MAX_BYTES + 1)
    if len(contents) > TEAM_SUBMISSION_MAX_BYTES:
        flash("Project submissions cannot exceed 25 MB.", "danger")
        return None
    if not contents:
        flash("The selected submission file is empty.", "danger")
        return None

    filename = f"{uuid.uuid4().hex}{extension}"
    submission_directory = os.path.join(
        current_app.config["UPLOAD_FOLDER"], "submissions"
    )
    os.makedirs(submission_directory, exist_ok=True)
    with open(os.path.join(submission_directory, filename), "wb") as file_stream:
        file_stream.write(contents)

    submission = TeamSubmission.query.filter_by(
        team_id=team.id, submission_type=submission_type
    ).first()
    previous_filename = submission.filename if submission else None
    if submission:
        submission.filename = filename
        submission.original_name = original_name
        submission.uploaded_at = datetime.utcnow()
    else:
        db.session.add(TeamSubmission(
            team_id=team.id,
            submission_type=submission_type,
            filename=filename,
            original_name=original_name,
        ))
    return filename, previous_filename


MAX_CERTIFICATE_ZIP_ENTRIES = 500
MAX_CERTIFICATE_PDF_SIZE = 25 * 1024 * 1024
MAX_CERTIFICATE_ZIP_SIZE = 300 * 1024 * 1024
MAX_CERTIFICATE_TEMPLATE_SIZE = 20 * 1024 * 1024
MAX_CERTIFICATE_DOCX_SIZE = 20 * 1024 * 1024
MAX_CERTIFICATE_CSV_SIZE = 5 * 1024 * 1024
MAX_CERTIFICATE_CSV_ROWS = 500
MAX_CERTIFICATE_SPREADSHEET_COLUMNS = 100


def certificate_excel_rows(file_data):
    try:
        workbook = load_workbook(
            io.BytesIO(file_data),
            read_only=True,
            data_only=True,
        )
    except (
        EOFError,
        InvalidFileException,
        KeyError,
        OSError,
        ValueError,
        zipfile.BadZipFile,
    ) as error:
        raise ValueError("Upload a valid .xlsx Excel workbook.") from error
    try:
        worksheet = workbook.active
        if (
            worksheet is None
            or worksheet.max_row is None
            or worksheet.max_row < 2
            or worksheet.max_row > MAX_CERTIFICATE_CSV_ROWS + 1
            or worksheet.max_column is None
            or worksheet.max_column > MAX_CERTIFICATE_SPREADSHEET_COLUMNS
        ):
            raise ValueError(
                "The first worksheet must have a header and 1–500 rows, "
                "with no more than 100 columns."
            )
        values = worksheet.iter_rows(values_only=True)
        raw_headers = next(values, None)
        if not raw_headers:
            raise ValueError("The Excel worksheet needs a header row.")
        headers = [str(value).strip() if value is not None else "" for value in raw_headers]
        if not headers or any(not header for header in headers):
            raise ValueError("Excel column headers must be non-empty.")
        if len(set(headers)) != len(headers):
            raise ValueError("Excel column headers must be unique after trimming spaces.")

        rows = []
        for row_number, values_row in enumerate(values, 2):
            row = {
                header: excel_cell_text(values_row[index] if index < len(values_row) else None)
                for index, header in enumerate(headers)
            }
            if not any(value.strip() for value in row.values()):
                continue
            rows.append(row)
            if len(rows) > MAX_CERTIFICATE_CSV_ROWS:
                raise ValueError("Excel must have no more than 500 recipient rows.")
        if not rows:
            raise ValueError("The first worksheet must contain at least one recipient row.")
        return headers, rows
    finally:
        workbook.close()


def excel_cell_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
CERTIFICATE_TEMPLATE_POSITIONS = {
    "recipient_name": {"x": 61, "y": 17, "size": 26},
    "event_name": {"x": 61, "y": 46, "size": 14},
    "event_date": {"x": 61, "y": 55, "size": 12},
    "event_venue": {"x": 61, "y": 61, "size": 12},
    "team_id": {"x": 61, "y": 70, "size": 14},
    "project_title": {"x": 61, "y": 80, "size": 12},
    "domain_name": {"x": 61, "y": 90, "size": 12},
    "certificate_number": {"x": 80, "y": 94, "size": 9},
    "qr": {"x": 10, "y": 90, "size": 9},
}


def new_certificate_id():
    return "FTI26-" + uuid.uuid4().hex.upper()


def certificate_verification_url(certificate_id):
    verification_path = url_for(
        "main.verify_certificate",
        certificate_id=certificate_id,
    )
    public_base_url = current_app.config.get("CERTIFICATE_PUBLIC_BASE_URL", "").strip()
    if public_base_url:
        return f"{public_base_url.rstrip('/')}{verification_path}"
    return url_for(
        "main.verify_certificate",
        certificate_id=certificate_id,
        _external=True,
    )


CERTIFICATE_TYPE_FIELDS = {
    "Certificate of Participation": (
        "recipient_name",
        "event_name",
        "domain_name",
        "team_id",
        "project_title",
        "event_date",
        "event_venue",
        "certificate_number",
        "qr",
    ),
    "Certificate of Winner": (
        "recipient_name",
        "domain_name",
        "event_name",
        "team_id",
        "project_title",
        "event_venue",
        "event_date",
        "certificate_number",
        "qr",
    ),
    "Certificate of Appreciation - Mentor": (
        "recipient_name",
        "event_name",
        "domain_name",
        "event_venue",
        "event_date",
        "certificate_number",
        "qr",
    ),
    "Certificate of Appreciation - Coordinator": (
        "recipient_name",
        "event_name",
        "domain_name",
        "event_venue",
        "event_date",
        "certificate_number",
        "qr",
    ),
    "Certificate of Appreciation - Club Head and Judges": (
        "recipient_name",
        "event_name",
        "event_date",
        "event_venue",
        "certificate_number",
        "qr",
    ),
}

CERTIFICATE_FIELD_LABELS = {
    "recipient_name": "Recipient name",
    "event_name": "Event name",
    "domain_name": "Domain name",
    "team_id": "Team ID",
    "project_title": "Project title",
    "event_date": "Event date",
    "event_venue": "Event venue",
    "certificate_number": "Certificate number",
    "qr": "QR code",
}

CERTIFICATE_TEMPLATE_ROLES = {
    "Certificate of Participation": {"team_leader", "team_member"},
    "Certificate of Winner": {"team_leader", "team_member"},
    "Certificate of Appreciation - Mentor": {"mentor"},
    "Certificate of Appreciation - Coordinator": {"coordinator"},
    "Certificate of Appreciation - Club Head and Judges": {"admin", "judge"},
}

def certificate_template_snapshot(template):
    return {
        "event_name": template.event_name,
        "event_date": template.event_date,
        "event_venue": template.event_venue,
        "recipient_role": template.recipient_role,
        "background_file": template.background_file,
        "document_file": template.document_file,
        "positions_json": template.positions_json,
        "page_width": template.page_width,
        "page_height": template.page_height,
    }


def has_certificate_word_source(certificate):
    if certificate.template_snapshot_json:
        try:
            snapshot = json.loads(certificate.template_snapshot_json)
        except (TypeError, json.JSONDecodeError):
            return False
        document_file = (
            snapshot.get("document_file")
            if isinstance(snapshot, dict)
            else None
        )
    else:
        document_file = (
            certificate.template.document_file if certificate.template else None
        )
    if not document_file or os.path.basename(document_file) != document_file:
        return False
    return os.path.isfile(os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "certificate_templates",
        document_file,
    ))


@main.app_context_processor
def inject_certificate_download_helpers():
    return {"has_certificate_word_source": has_certificate_word_source}


def editable_certificate_docx(type_name):
    fields = CERTIFICATE_TYPE_FIELDS[type_name]
    labels = {
        "recipient_name": "Name",
        "event_name": "Event",
        "event_date": "Date",
        "event_venue": "Venue",
        "domain_name": "Domain",
        "team_id": "Team ID",
        "project_title": "Project",
        "certificate_number": "Certificate number",
        "qr": "QR code",
    }
    paragraphs = [
        ("FUTURE-TECH INNOVISION 2026", 28, True),
        (type_name, 22, True),
        ("This certificate is presented to", 14, False),
        ("<name>", 24, True),
        ("for participation and contribution to the exhibition.", 12, False),
    ]
    paragraphs.extend(
        (f"{labels[field]}: <{field}>", 12, False)
        for field in fields
        if field not in {"recipient_name", "event_name", "qr"}
    )
    paragraphs.extend(
        [
            ("Event: <event_name>", 12, False),
            ("<qr>", 10, False),
        ]
    )
    document_paragraphs = []
    for text, size, bold in paragraphs:
        bold_run = "<w:b/>" if bold else ""
        document_paragraphs.append(
            "<w:p><w:pPr><w:jc w:val=\"center\"/>"
            "<w:spacing w:after=\"240\"/></w:pPr><w:r><w:rPr>"
            f"<w:sz w:val=\"{size * 2}\"/>{bold_run}</w:rPr>"
            f"<w:t xml:space=\"preserve\">{escape(text)}</w:t>"
            "</w:r></w:p>"
        )
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>"
        + "".join(document_paragraphs)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900" '
        'w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    package_relationships = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", package_relationships)
        archive.writestr("word/document.xml", document_xml)
    output.seek(0)
    return output


def serialize_opc_xml_with_default_namespace(element, namespace):
    serialized = ET.tostring(element, encoding="utf-8", xml_declaration=True)
    root_start = serialized.find(b"<", serialized.find(b"?>") + 2)
    root_end = serialized.find(b">", root_start)
    root_tag = serialized[root_start:root_end]
    prefix_match = re.match(rb"<([A-Za-z_][\w.-]*):", root_tag)
    namespace_bytes = namespace.encode("utf-8")
    if prefix_match is None:
        if b'xmlns="' + namespace_bytes + b'"' in root_tag:
            return serialized
        raise ValueError("The Open Packaging Convention XML namespace is missing.")

    prefix = prefix_match.group(1)
    declaration = re.search(
        rb'\s+xmlns:' + re.escape(prefix) + rb'="([^"]+)"',
        root_tag,
    )
    if declaration is None or declaration.group(1) != namespace_bytes:
        raise ValueError("The Open Packaging Convention XML namespace is invalid.")

    serialized = serialized.replace(
        declaration.group(0),
        b' xmlns="' + namespace_bytes + b'"',
        1,
    )
    serialized = serialized.replace(b"<" + prefix + b":", b"<")
    return serialized.replace(b"</" + prefix + b":", b"</")


def merged_certificate_docx(
    template_path, values, verification_url, template_positions=None
):
    try:
        with zipfile.ZipFile(template_path) as source:
            entries = source.infolist()
            if len(entries) > 2000 or sum(entry.file_size for entry in entries) > MAX_CERTIFICATE_DOCX_SIZE * 4:
                abort(503, description="The saved Word certificate template is too large to merge.")
            parts = {info.filename: source.read(info.filename) for info in entries}
    except (OSError, zipfile.BadZipFile, RuntimeError) as error:
        abort(503, description="The saved Word certificate template could not be read.")

    document_name = "word/document.xml"
    if document_name not in parts:
        abort(503, description="The Word template does not contain a document body.")
    docx_namespaces = {}
    for _, namespace in ET.iterparse(
        io.BytesIO(parts[document_name]), events=("start-ns",)
    ):
        prefix, uri = namespace
        docx_namespaces[prefix] = uri
        if prefix != "xml" and not re.fullmatch(r"ns\d+", prefix):
            ET.register_namespace(prefix, uri)
    try:
        root = ET.fromstring(parts[document_name])
    except ET.ParseError:
        abort(503, description="The saved Word certificate document is invalid.")

    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ns = {
        "w": word_ns,
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
        "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    }
    ET.register_namespace("w", ns["w"])
    ET.register_namespace("r", ns["r"])
    aliases = {
        "name": "recipient_name",
        "role": "recipient_role",
        "event": "event_name",
        "day": "event_day",
        "date": "event_date",
        "venue": "event_venue",
        "domain": "domain_name",
        "team": "team_id",
        "title": "project_title",
        "certificate_number": "certificate_id",
        "certificate_no": "certificate_id",
        "qr": "qr",
        "verification_url": "verification_url",
    }
    merge_values = {**values, "verification_url": verification_url}
    template_positions = template_positions or {}
    field_mappings = template_positions.get("docx_field_mappings", {})
    text_mappings = template_positions.get("docx_text_mappings", [])
    field_styles = template_positions.get("docx_field_styles", {})
    token_pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}|<([a-zA-Z0-9_]+)>")
    qr_image_needed = False
    qr_target_paragraph = None
    certificate_number_paragraphs = set()

    def apply_run_style(run, style):
        if not style:
            return
        run_properties = run.find("w:rPr", ns)
        if run_properties is None:
            run_properties = ET.Element(f"{{{word_ns}}}rPr")
            run.insert(0, run_properties)
        for tag, enabled in (("b", style["bold"]), ("i", style["italic"])):
            node = run_properties.find(f"w:{tag}", ns)
            if node is None:
                node = ET.SubElement(run_properties, f"{{{word_ns}}}{tag}")
            node.set(f"{{{word_ns}}}val", "1" if enabled else "0")
        color_node = run_properties.find("w:color", ns)
        if color_node is None:
            color_node = ET.SubElement(run_properties, f"{{{word_ns}}}color")
        color_node.set(f"{{{word_ns}}}val", style["color"].lstrip("#").upper())
        font_node = run_properties.find("w:rFonts", ns)
        if font_node is None:
            font_node = ET.SubElement(run_properties, f"{{{word_ns}}}rFonts")
        font_node.set(f"{{{word_ns}}}ascii", style["font"])
        font_node.set(f"{{{word_ns}}}hAnsi", style["font"])
        size_node = run_properties.find("w:sz", ns)
        if size_node is None:
            size_node = ET.SubElement(run_properties, f"{{{word_ns}}}sz")
        size_node.set(f"{{{word_ns}}}val", str(int(round(style["size"] * 2))))

    parent_lookup = {
        child: parent
        for parent in root.iter()
        for child in list(parent)
    }
    paragraphs = root.findall(".//w:p", ns)
    for paragraph_index, paragraph in enumerate(paragraphs):
        text_nodes = paragraph.findall(".//w:t", ns)
        if not text_nodes:
            continue
        paragraph_text = "".join(node.text or "" for node in text_nodes)
        replacements = []
        for match in token_pattern.finditer(paragraph_text):
            token = (match.group(1) or match.group(2)).casefold()
            field = field_mappings.get(token, aliases.get(token, token))
            replacements.append((match.start(), match.end(), token, field, False))
        for mapping in text_mappings:
            if mapping["paragraph_index"] != paragraph_index:
                continue
            for match in re.finditer(re.escape(mapping["text"]), paragraph_text):
                replacements.append((
                    match.start(),
                    match.end(),
                    mapping["field"],
                    mapping["field"],
                    True,
                ))
        if not replacements:
            continue

        replacements.sort(key=lambda item: (item[0], -(item[1] - item[0])))
        segments = []
        cursor = 0
        for start, end, token, field, is_text_mapping in replacements:
            if start < cursor:
                continue
            if start > cursor:
                segments.append((paragraph_text[cursor:start], None))
            if field == "qr":
                qr_image_needed = True
                if qr_target_paragraph is None:
                    qr_target_paragraph = paragraph
            elif (
                field in {"certificate_id", "certificate_number"}
                and not is_text_mapping
            ):
                certificate_number_paragraphs.add(paragraph)
            elif field in {"certificate_id", "certificate_number"}:
                replacement = str(merge_values.get(field, ""))
                if replacement:
                    segments.append((
                        replacement,
                        field_styles.get(token, field_styles.get(field)),
                    ))
            else:
                replacement = str(merge_values.get(
                    field, paragraph_text[start:end]
                ))
                if replacement:
                    segments.append((
                        replacement,
                        field_styles.get(token, field_styles.get(field)),
                    ))
            cursor = end
        if cursor < len(paragraph_text):
            segments.append((paragraph_text[cursor:], None))

        if not segments:
            for node in text_nodes:
                node.text = ""
            continue
        first_node = text_nodes[0]
        first_run = parent_lookup.get(first_node)
        while first_run is not None and first_run.tag != f"{{{word_ns}}}r":
            first_run = parent_lookup.get(first_run)
        if first_run is None:
            first_node.text = "".join(segment for segment, _ in segments)
            for node in text_nodes[1:]:
                node.text = ""
            continue
        source_properties = first_run.find("w:rPr", ns)
        base_properties = (
            copy.deepcopy(source_properties)
            if source_properties is not None
            else None
        )
        first_node.text = segments[0][0]
        if segments[0][1]:
            apply_run_style(first_run, segments[0][1])
        for node in text_nodes[1:]:
            node.text = ""
        insertion_parent = first_run
        while (
            parent_lookup.get(insertion_parent) is not paragraph
            and parent_lookup.get(insertion_parent) is not None
        ):
            insertion_parent = parent_lookup[insertion_parent]
        insert_at = list(paragraph).index(insertion_parent) + 1
        for segment_text, style in segments[1:]:
            new_run = ET.Element(f"{{{word_ns}}}r")
            if style:
                properties = (
                    copy.deepcopy(base_properties)
                    if base_properties is not None
                    else ET.Element(f"{{{word_ns}}}rPr")
                )
                new_run.append(properties)
                apply_run_style(new_run, style)
            elif base_properties is not None:
                new_run.append(copy.deepcopy(base_properties))
            new_text = ET.SubElement(new_run, f"{{{word_ns}}}t")
            new_text.text = segment_text
            if segment_text != segment_text.strip():
                new_text.set(
                    "{http://www.w3.org/XML/1998/namespace}space",
                    "preserve",
                )
            paragraph.insert(insert_at, new_run)
            insert_at += 1
    for paragraph in certificate_number_paragraphs:
        parent = parent_lookup.get(paragraph)
        if parent is not None:
            parent.remove(paragraph)

    relationships_name = "word/_rels/document.xml.rels"
    if relationships_name in parts:
        try:
            relationships = ET.fromstring(parts[relationships_name])
        except ET.ParseError:
            abort(503, description="The Word template relationships are invalid.")
    else:
        relationships = ET.Element(
            f"{{{ns['rel']}}}Relationships"
        )

    relationship_ids = {
        element.get("Id", "")
        for element in relationships.findall("rel:Relationship", ns)
    }
    next_id = 1
    while f"rId{next_id}" in relationship_ids:
        next_id += 1
    image_relationship_id = f"rId{next_id}"
    image_name = f"certificate-qr-{uuid.uuid4().hex}.png"
    qr = qrcode.make(verification_url).convert("RGB")
    certificate_id = str(merge_values["certificate_id"])
    try:
        label_font = ImageFont.load_default(size=28)
        number_font = ImageFont.load_default(size=32)
    except TypeError:
        label_font = ImageFont.load_default()
        number_font = ImageFont.load_default()
    qr_draw = ImageDraw.Draw(qr)
    label = "Certificate No."
    label_bounds = qr_draw.textbbox((0, 0), label, font=label_font)
    number_bounds = qr_draw.textbbox((0, 0), certificate_id, font=number_font)
    label_width = label_bounds[2] - label_bounds[0]
    number_width = number_bounds[2] - number_bounds[0]
    composition_width = max(qr.width * 2.5, label_width + 24, number_width + 24)
    padding = 8
    label_height = label_bounds[3] - label_bounds[1]
    number_height = number_bounds[3] - number_bounds[1]
    composition_height = qr.height + padding * 3 + label_height + number_height
    qr_composition = Image.new(
        "RGB",
        (int(math.ceil(composition_width)), composition_height),
        "white",
    )
    qr_composition.paste(qr, ((qr_composition.width - qr.width) // 2, 0))
    text_draw = ImageDraw.Draw(qr_composition)
    label_x = (qr_composition.width - label_width) / 2 - label_bounds[0]
    label_y = qr.height + padding - label_bounds[1]
    text_draw.text((label_x, label_y), label, font=label_font, fill="black")
    number_x = (qr_composition.width - number_width) / 2 - number_bounds[0]
    number_y = qr.height + padding * 2 + label_height - number_bounds[1]
    text_draw.text(
        (number_x, number_y),
        certificate_id,
        font=number_font,
        fill="black",
    )
    qr_buffer = io.BytesIO()
    qr_composition.save(qr_buffer, format="PNG")

    word_drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    drawing_ids = [
        int(item.get("id", "0"))
        for item in root.findall(".//wp:docPr", {"wp": word_drawing_ns})
        if item.get("id", "").isdigit()
    ]
    drawing_id = max(drawing_ids, default=0) + 1
    body = root.find("w:body", ns)
    if body is None:
        abort(503, description="The Word template has no document body.")
    qr_position = template_positions.get(
        "qr", CERTIFICATE_TEMPLATE_POSITIONS["qr"]
    )
    if not isinstance(qr_position, dict):
        abort(503, description="The saved Word QR position is invalid.")
    try:
        qr_x = float(qr_position["x"])
        qr_y = float(qr_position["y"])
        qr_size = float(qr_position["size"])
    except (KeyError, TypeError, ValueError):
        abort(503, description="The saved Word QR position is invalid.")
    if (
        not all(math.isfinite(value) for value in (qr_x, qr_y, qr_size))
        or not 0 <= qr_x <= 100
        or not 0 <= qr_y <= 100
        or not 3 <= qr_size <= 20
    ):
        abort(503, description="The saved Word QR position is invalid.")
    section_page_sizes = root.findall(".//w:sectPr/w:pgSz", ns)
    page_size = section_page_sizes[-1] if section_page_sizes else None
    try:
        page_width_twips = int(
            page_size.get(f"{{{word_ns}}}w", "11906")
            if page_size is not None
            else 11906
        )
        page_height_twips = int(
            page_size.get(f"{{{word_ns}}}h", "16838")
            if page_size is not None
            else 16838
        )
    except (TypeError, ValueError):
        abort(503, description="The Word template page size is invalid.")
    if page_width_twips <= 0 or page_height_twips <= 0:
        abort(503, description="The Word template page size is invalid.")
    page_width_emu = page_width_twips * 635
    page_height_emu = page_height_twips * 635
    qr_size_emu = int(page_width_emu * qr_size / 100)
    composition_width_emu = int(
        qr_size_emu * qr_composition.width / qr.width
    )
    composition_height_emu = int(
        qr_size_emu * qr_composition.height / qr.height
    )
    qr_center_x_emu = int(page_width_emu * qr_x / 100)
    qr_top_emu = int(page_height_emu * qr_y / 100 - qr_size_emu / 2)
    qr_left_emu = int(qr_center_x_emu - composition_width_emu / 2)
    qr_left_emu = max(0, min(page_width_emu - composition_width_emu, qr_left_emu))
    qr_top_emu = max(0, min(page_height_emu - composition_height_emu, qr_top_emu))
    ET.SubElement(
        relationships,
        f"{{{ns['rel']}}}Relationship",
        {
            "Id": image_relationship_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "Target": f"media/{image_name}",
        },
    )
    parts[relationships_name] = serialize_opc_xml_with_default_namespace(
        relationships, ns["rel"]
    )
    content_types_name = "[Content_Types].xml"
    try:
        content_types = ET.fromstring(parts[content_types_name])
    except (KeyError, ET.ParseError):
        abort(503, description="The Word template content types are invalid.")
    if not any(
        item.get("Extension", "").casefold() == "png"
        for item in content_types.findall("ct:Default", ns)
    ):
        ET.SubElement(
            content_types,
            f"{{{ns['ct']}}}Default",
            {
                "Extension": "png",
                "ContentType": "image/png",
            },
        )
    parts[content_types_name] = serialize_opc_xml_with_default_namespace(
        content_types, ns["ct"]
    )

    drawing_xml = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" '
        'simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
        'layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{qr_left_emu}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{qr_top_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{composition_width_emu}" cy="{composition_height_emu}"/>'
        '<wp:wrapNone/>'
        f'<wp:docPr id="{drawing_id}" name="Certificate verification QR"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr><a:graphic><a:graphicData '
        'uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Certificate QR"/>'
        '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
        f'<a:blip r:embed="{image_relationship_id}"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{composition_width_emu}" cy="{composition_height_emu}"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr></pic:pic></a:graphicData></a:graphic>'
        '</wp:anchor></w:drawing></w:r>'
    )
    drawing = ET.fromstring(drawing_xml)
    if qr_image_needed:
        target_paragraph = qr_target_paragraph
        if target_paragraph is None:
            target_paragraph = ET.Element(f"{{{word_ns}}}p")
            section_properties = body.find("w:sectPr", ns)
            if section_properties is None:
                body.append(target_paragraph)
            else:
                body.insert(list(body).index(section_properties), target_paragraph)
    else:
        target_paragraph = ET.Element(f"{{{word_ns}}}p")
        section_properties = body.find("w:sectPr", ns)
        if section_properties is None:
            body.append(target_paragraph)
        else:
            body.insert(list(body).index(section_properties), target_paragraph)
    if target_paragraph.find("w:pPr", ns) is None:
        paragraph_properties = ET.Element(f"{{{word_ns}}}pPr")
        target_paragraph.insert(0, paragraph_properties)
    else:
        paragraph_properties = target_paragraph.find("w:pPr", ns)
    alignment = paragraph_properties.find("w:jc", ns)
    if alignment is None:
        alignment = ET.SubElement(paragraph_properties, f"{{{word_ns}}}jc")
    alignment.set(f"{{{word_ns}}}val", "left")
    target_paragraph.append(drawing)

    serialized_document = ET.tostring(
        root, encoding="utf-8", xml_declaration=True
    ).decode("utf-8")
    root_start = serialized_document.find("<", serialized_document.find("?>") + 2)
    root_end = serialized_document.find(">", root_start)
    root_tag = serialized_document[:root_end]
    declarations = []
    for prefix, uri in docx_namespaces.items():
        declaration_name = "xmlns" if not prefix else f"xmlns:{prefix}"
        if f"{declaration_name}=" in root_tag:
            continue
        escaped_uri = uri.replace("&", "&amp;").replace('"', "&quot;")
        declarations.append(f' {declaration_name}="{escaped_uri}"')
    if declarations:
        serialized_document = (
            serialized_document[:root_end]
            + "".join(declarations)
            + serialized_document[root_end:]
        )
    parts[document_name] = serialized_document.encode("utf-8")
    parts[f"word/media/{image_name}"] = qr_buffer.getvalue()
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as destination:
        for filename, contents in parts.items():
            destination.writestr(filename, contents)
    output.seek(0)
    return output


def convert_certificate_docx_to_pdf(document_data):
    converter = shutil.which("libreoffice") or shutil.which("soffice")
    if not converter:
        raise RuntimeError(
            "Word-to-PDF conversion is unavailable. The server requires LibreOffice."
        )
    with tempfile.TemporaryDirectory(prefix="certificate-docx-") as temp_dir:
        input_path = os.path.join(temp_dir, "certificate.docx")
        with open(input_path, "xb") as source:
            source.write(document_data)
        output_dir = os.path.join(temp_dir, "output")
        profile_dir = os.path.join(temp_dir, "profile")
        os.makedirs(output_dir)
        profile_url = Path(profile_dir).as_uri()
        try:
            result = subprocess.run(
                [
                    converter,
                    f"-env:UserInstallation={profile_url}",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    output_dir,
                    input_path,
                ],
                check=False,
                capture_output=True,
                timeout=60,
            )
        except subprocess.TimeoutExpired as error:
            raise RuntimeError(
                "Word-to-PDF conversion exceeded the 60-second limit."
            ) from error
        pdf_path = os.path.join(output_dir, "certificate.pdf")
        if result.returncode or not os.path.isfile(pdf_path):
            detail = (result.stderr or result.stdout).decode(
                "utf-8", errors="replace"
            ).strip()
            raise RuntimeError(
                "LibreOffice could not render the Word certificate."
                + (f" {detail[:300]}" if detail else "")
            )
        with open(pdf_path, "rb") as rendered:
            pdf_data = rendered.read()
    try:
        with pymupdf.open(stream=pdf_data, filetype="pdf") as document:
            if len(document) != 1:
                raise ValueError(
                    "A certificate Word template must render to exactly one PDF page."
                )
    except (pymupdf.FileDataError, RuntimeError) as error:
        raise ValueError("The Word certificate did not produce a valid PDF.") from error
    return pdf_data


CERTIFICATE_CSV_COLUMN_ALIASES = {
    "username_column": ("username", "accountusername", "studentid", "login"),
    "recipient_name_column": ("name", "recipientname", "fullname"),
    "domain_column": ("domainname", "domain"),
    "team_id_column": ("teamid", "teamcode", "team"),
    "team_code_column": ("teamid", "teamcode", "team"),
    "project_title_column": ("projecttitle", "project", "title"),
    "event_name_column": ("eventname", "event"),
    "event_date_column": ("eventdate", "date"),
    "event_venue_column": ("eventvenue", "venue", "location"),
    "recipient_role_column": ("role", "recipientrole"),
    "event_day_column": ("day", "eventday"),
}


def parse_certificate_positions(raw):
    try:
        supplied = json.loads(raw) if raw else {}
    except json.JSONDecodeError:
        abort(400, description="Certificate field positions are invalid.")
    if not isinstance(supplied, dict):
        abort(400, description="Certificate field positions must be an object.")
    docx_pdf_primary = supplied.get("docx_pdf_primary", False)
    if not isinstance(docx_pdf_primary, bool):
        abort(400, description="The Word PDF rendering option is invalid.")
    positions = {}
    for field, default in CERTIFICATE_TEMPLATE_POSITIONS.items():
        legacy_field = "team_code" if field == "team_id" else field
        value = supplied.get(field, supplied.get(legacy_field, default))
        if not isinstance(value, dict):
            abort(400, description=f"Invalid position for {field}.")
        try:
            x = float(value.get("x", default["x"]))
            y = float(value.get("y", default["y"]))
            size = float(value.get("size", default["size"]))
        except (TypeError, ValueError):
            abort(400, description=f"Invalid position for {field}.")
        max_size = 20 if field == "qr" else 48
        min_size = 3 if field == "qr" else 6
        if (
            not all(math.isfinite(number) for number in (x, y, size))
            or not 0 <= x <= 100
            or not 0 <= y <= 100
            or not min_size <= size <= max_size
        ):
            abort(400, description=f"Position or size for {field} is out of range.")
        font = value.get("font", "Helvetica")
        alignment = value.get("align", "center")
        color = value.get("color", "#20242c")
        bold = value.get("bold", field == "recipient_name")
        italic = value.get("italic", False)
        if (
            not isinstance(font, str)
            or font not in {"Helvetica", "Times", "Courier"}
            or not isinstance(alignment, str)
            or alignment not in {"left", "center", "right"}
            or not isinstance(color, str)
            or not re.fullmatch(r"#[0-9a-fA-F]{6}", color)
            or not isinstance(bold, bool)
            or not isinstance(italic, bool)
        ):
            abort(400, description=f"Invalid text styling for certificate field {field}.")
        positions[field] = {
            "x": x,
            "y": y,
            "size": size,
            "font": font,
            "align": alignment,
            "color": color.lower(),
            "bold": bold,
            "italic": italic,
        }
    field_mappings = supplied.get("docx_field_mappings", {})
    if not isinstance(field_mappings, dict) or len(field_mappings) > 100:
        abort(400, description="Word field mappings must be an object with no more than 100 entries.")
    allowed_docx_fields = {
        "recipient_name",
        "recipient_role",
        "domain_name",
        "event_name",
        "event_day",
        "event_date",
        "event_venue",
        "team_id",
        "project_title",
        "certificate_id",
        "verification_url",
    }
    normalized_mappings = {}
    for token, field in field_mappings.items():
        normalized_token = str(token).strip().casefold()
        if (
            not re.fullmatch(r"[a-z0-9_]{1,60}", normalized_token)
            or not isinstance(field, str)
            or field not in allowed_docx_fields
        ):
            abort(400, description="A Word field mapping contains an unsupported token or field.")
        normalized_mappings[normalized_token] = field
    positions["docx_field_mappings"] = normalized_mappings
    raw_text_mappings = supplied.get("docx_text_mappings", [])
    if not isinstance(raw_text_mappings, list) or len(raw_text_mappings) > 100:
        abort(400, description="Word text mappings must contain no more than 100 entries.")
    normalized_text_mappings = []
    seen_text_mappings = set()
    for mapping in raw_text_mappings:
        if not isinstance(mapping, dict):
            abort(400, description="A Word text mapping is invalid.")
        paragraph_index = mapping.get("paragraph_index")
        text = mapping.get("text")
        field = mapping.get("field")
        if (
            not isinstance(paragraph_index, int)
            or isinstance(paragraph_index, bool)
            or not 0 <= paragraph_index < 10000
            or not isinstance(text, str)
            or not text.strip()
            or len(text) > 200
            or not isinstance(field, str)
            or field not in allowed_docx_fields
        ):
            abort(400, description="A Word text mapping has an unsupported selection or field.")
        key = (paragraph_index, text)
        if key in seen_text_mappings:
            abort(400, description="A Word text selection can only be mapped once.")
        seen_text_mappings.add(key)
        normalized_text_mappings.append({
            "paragraph_index": paragraph_index,
            "text": text,
            "field": field,
        })
    positions["docx_text_mappings"] = normalized_text_mappings
    raw_field_styles = supplied.get("docx_field_styles", {})
    if not isinstance(raw_field_styles, dict) or len(raw_field_styles) > 100:
        abort(400, description="Word field styles must be an object with no more than 100 entries.")
    normalized_field_styles = {}
    for token, style in raw_field_styles.items():
        normalized_token = str(token).strip().casefold()
        if (
            not re.fullmatch(r"[a-z0-9_]{1,60}", normalized_token)
            or not isinstance(style, dict)
        ):
            abort(400, description="A Word field style contains an invalid placeholder.")
        font = style.get("font", "Arial")
        color = style.get("color", "#20242c")
        bold = style.get("bold", False)
        italic = style.get("italic", False)
        try:
            size = float(style.get("size", 12))
        except (TypeError, ValueError):
            abort(400, description="A Word field font size is invalid.")
        if (
            not isinstance(font, str)
            or font not in {"Arial", "Times New Roman", "Courier New"}
            or not isinstance(color, str)
            or not re.fullmatch(r"#[0-9a-fA-F]{6}", color)
            or not isinstance(bold, bool)
            or not isinstance(italic, bool)
            or not math.isfinite(size)
            or not 6 <= size <= 48
        ):
            abort(400, description="A Word field style contains an unsupported font, color, or size.")
        normalized_field_styles[normalized_token] = {
            "font": font,
            "color": color.lower(),
            "bold": bold,
            "italic": italic,
            "size": size,
        }
    positions["docx_field_styles"] = normalized_field_styles
    custom_texts = supplied.get("custom_texts", [])
    if not isinstance(custom_texts, list) or len(custom_texts) > 40:
        abort(400, description="Add no more than 40 custom certificate text elements.")
    normalized_custom_texts = []
    seen_ids = set()
    allowed_fonts = {"Helvetica", "Times", "Courier"}
    allowed_alignments = {"left", "center", "right"}
    for item in custom_texts:
        if not isinstance(item, dict):
            abort(400, description="Invalid custom certificate text element.")
        element_id = str(item.get("id", ""))
        text = item.get("text", "")
        font = item.get("font", "Helvetica")
        alignment = item.get("align", "center")
        color = item.get("color", "#20242c")
        try:
            x = float(item.get("x", 50))
            y = float(item.get("y", 50))
            size = float(item.get("size", 24))
        except (TypeError, ValueError):
            abort(400, description="Invalid custom certificate text position or size.")
        if (
            not element_id
            or len(element_id) > 80
            or element_id in seen_ids
            or not isinstance(text, str)
            or not text.strip()
            or len(text) > 500
            or not isinstance(font, str)
            or font not in allowed_fonts
            or not isinstance(alignment, str)
            or alignment not in allowed_alignments
            or not isinstance(color, str)
            or not re.fullmatch(r"#[0-9A-Fa-f]{6}", color)
            or not all(math.isfinite(number) for number in (x, y, size))
            or not 0 <= x <= 100
            or not 0 <= y <= 100
            or not 6 <= size <= 100
            or not isinstance(item.get("bold", False), bool)
            or not isinstance(item.get("italic", False), bool)
        ):
            abort(400, description="Custom certificate text has invalid content or styling.")
        seen_ids.add(element_id)
        normalized_custom_texts.append({
            "id": element_id,
            "text": text,
            "x": x,
            "y": y,
            "size": size,
            "color": color.lower(),
            "font": font,
            "bold": item.get("bold", False),
            "italic": item.get("italic", False),
            "align": alignment,
        })
    positions["custom_texts"] = normalized_custom_texts
    cover_zones = supplied.get("cover_zones", [])
    if not isinstance(cover_zones, list) or len(cover_zones) > 20:
        abort(400, description="Add no more than 20 certificate cover areas.")
    normalized_cover_zones = []
    for zone in cover_zones:
        if not isinstance(zone, dict):
            abort(400, description="Invalid certificate cover area.")
        try:
            x = float(zone.get("x", 50))
            y = float(zone.get("y", 50))
            zone_width = float(zone.get("width", 20))
            zone_height = float(zone.get("height", 6))
        except (TypeError, ValueError):
            abort(400, description="Invalid certificate cover area dimensions.")
        color = zone.get("color", "#ffffff")
        if (
            not all(math.isfinite(number) for number in (x, y, zone_width, zone_height))
            or not 0 <= x <= 100
            or not 0 <= y <= 100
            or not 0.1 <= zone_width <= 100
            or not 0.1 <= zone_height <= 100
            or not isinstance(color, str)
            or not re.fullmatch(r"#[0-9A-Fa-f]{6}", color)
        ):
            abort(400, description="Certificate cover area has invalid size, position, or color.")
        normalized_cover_zones.append({
            "x": x,
            "y": y,
            "width": zone_width,
            "height": zone_height,
            "color": color.lower(),
        })
    positions["cover_zones"] = normalized_cover_zones
    positions["docx_pdf_primary"] = docx_pdf_primary
    return positions


def read_certificate_template_upload(upload):
    if not upload or not upload.filename:
        abort(400, description="Upload a Canva template as a one-page PDF or image.")
    extension = os.path.splitext(secure_filename(upload.filename))[1].lower()
    if extension not in {".pdf", ".png", ".jpg", ".jpeg"}:
        abort(400, description="Templates must be a one-page PDF, PNG, or JPEG.")
    file_data = upload.stream.read(MAX_CERTIFICATE_TEMPLATE_SIZE + 1)
    if not file_data or len(file_data) > MAX_CERTIFICATE_TEMPLATE_SIZE:
        abort(400, description="Certificate template must be no larger than 20 MB.")
    if extension == ".pdf":
        if not file_data.startswith(b"%PDF-"):
            abort(400, description="The uploaded certificate template is not a valid PDF.")
        with pymupdf.open(stream=file_data, filetype="pdf") as source:
            if len(source) != 1:
                abort(400, description="Use a one-page PDF certificate template.")
    elif extension == ".png":
        if not file_data.startswith(b"\x89PNG\r\n\x1a\n"):
            abort(400, description="The uploaded file is not a valid PNG image.")
    elif not file_data.startswith(b"\xff\xd8\xff"):
        abort(400, description="The uploaded file is not a valid JPEG image.")
    return extension, file_data


def read_certificate_template_document_upload(upload):
    if not upload or not upload.filename:
        return None
    extension = os.path.splitext(secure_filename(upload.filename))[1].lower()
    if extension != ".docx":
        abort(400, description="The editable certificate source must be a DOCX file.")
    file_data = upload.stream.read(MAX_CERTIFICATE_DOCX_SIZE + 1)
    if not file_data or len(file_data) > MAX_CERTIFICATE_DOCX_SIZE:
        abort(400, description="The editable certificate DOCX must be no larger than 20 MB.")
    try:
        with zipfile.ZipFile(io.BytesIO(file_data)) as archive:
            document_info = archive.getinfo("word/document.xml")
            archive.getinfo("[Content_Types].xml")
            entries = archive.infolist()
            if (
                len(entries) > 2000
                or sum(entry.file_size for entry in entries) > MAX_CERTIFICATE_DOCX_SIZE * 4
                or document_info.file_size > MAX_CERTIFICATE_DOCX_SIZE
            ):
                abort(400, description="The editable certificate DOCX contains oversized content.")
            ET.fromstring(archive.read(document_info))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError):
        abort(400, description="Upload a valid Microsoft Word DOCX document.")
    return file_data


def certificate_template_dimensions(extension, file_data):
    if extension == ".pdf":
        with pymupdf.open(stream=file_data, filetype="pdf") as source:
            page = source[0]
            return float(page.rect.width), float(page.rect.height)
    width, height = ImageReader(io.BytesIO(file_data)).getSize()
    scale = 842 / max(width, height)
    return float(width * scale), float(height * scale)


def certificate_template_background(path, width, height):
    if path.lower().endswith(".pdf"):
        with pymupdf.open(path) as source:
            if not len(source):
                abort(503, description="The approved certificate template is empty.")
            page = source[0]
            rect = page.rect
            pixmap = page.get_pixmap(
                matrix=pymupdf.Matrix(width / rect.width, height / rect.height),
                alpha=False,
            )
            return ImageReader(io.BytesIO(pixmap.tobytes("png")))
    return ImageReader(path)


def certificate_word_merge_data(certificate, template_values):
    document_file = template_values.get("document_file")
    if not document_file or os.path.basename(document_file) != document_file:
        abort(404, description="This certificate has no saved Word template.")
    document_path = os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "certificate_templates",
        document_file,
    )
    if not os.path.isfile(document_path):
        abort(503, description="The saved Word certificate template file is missing.")
    try:
        values = json.loads(certificate.fields_json or "{}")
        positions = json.loads(template_values.get("positions_json", "{}") or "{}")
    except (TypeError, json.JSONDecodeError):
        abort(503, description="The saved Word certificate data is invalid.")
    if not isinstance(values, dict) or not isinstance(positions, dict):
        abort(503, description="The saved Word certificate data is invalid.")
    values.update({
        "recipient_name": certificate.recipient_name,
        "recipient_role": certificate.recipient_role,
        "certificate_id": certificate.certificate_id,
        "certificate_number": certificate.certificate_id,
        "team_id": certificate.team_code,
    })
    return document_path, values, positions


def render_certificate_word_pdf(certificate, template_values):
    document_path, values, positions = certificate_word_merge_data(
        certificate, template_values
    )
    verification_url = certificate_verification_url(certificate.certificate_id)
    merged_document = merged_certificate_docx(
        document_path,
        values,
        verification_url,
        positions,
    )
    try:
        return io.BytesIO(
            convert_certificate_docx_to_pdf(merged_document.getvalue())
        )
    except (RuntimeError, ValueError) as error:
        abort(503, description=str(error))


def render_template_certificate(certificate):
    current_template = certificate.template
    has_snapshot = bool(certificate.template_snapshot_json)
    if (
        (not current_template and not has_snapshot)
        or (current_template and not current_template.approved and not has_snapshot)
    ):
        abort(503, description="This certificate template has not been approved.")
    try:
        template_values = (
            json.loads(certificate.template_snapshot_json)
            if has_snapshot
            else certificate_template_snapshot(current_template)
        )
    except (TypeError, json.JSONDecodeError):
        abort(503, description="Certificate template snapshot is invalid.")
    if not isinstance(template_values, dict):
        abort(503, description="Certificate template snapshot is invalid.")
    if template_values.get("document_file"):
        saved_positions = parse_certificate_positions(
            template_values.get("positions_json", "")
        )
        if saved_positions["docx_pdf_primary"]:
            return render_certificate_word_pdf(certificate, template_values)
    background_file = template_values.get("background_file", "")
    if os.path.basename(background_file) != background_file:
        abort(503, description="The certificate template file path is invalid.")
    template_path = os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "certificate_templates",
        background_file,
    )
    if not os.path.isfile(template_path):
        abort(503, description="The approved certificate template file is missing.")

    width = template_values.get("page_width") or 842
    height = template_values.get("page_height") or 595
    output = io.BytesIO()
    document = canvas.Canvas(output, pagesize=(width, height))
    document.drawImage(
        certificate_template_background(template_path, width, height),
        0,
        0,
        width=width,
        height=height,
    )
    try:
        positions = parse_certificate_positions(template_values.get("positions_json", ""))
        values = json.loads(certificate.fields_json or "{}")
    except (TypeError, json.JSONDecodeError):
        abort(503, description="Certificate template data is invalid.")
    if not isinstance(values, dict):
        abort(503, description="Certificate recipient data is invalid.")
    for zone in positions["cover_zones"]:
        zone_width = width * zone["width"] / 100
        zone_height = height * zone["height"] / 100
        center_x = width * zone["x"] / 100
        center_y = height * (1 - zone["y"] / 100)
        document.setFillColor(colors.HexColor(zone["color"]))
        document.rect(
            center_x - zone_width / 2,
            center_y - zone_height / 2,
            zone_width,
            zone_height,
            fill=1,
            stroke=0,
        )
    values = {
        "recipient_name": certificate.recipient_name,
        "certificate_number": certificate.certificate_id,
        "certificate_id": certificate.certificate_id,
        "event_name": template_values.get("event_name", ""),
        "event_date": template_values.get("event_date", ""),
        "event_venue": template_values.get("event_venue", ""),
        "team_id": certificate.team_code,
        **values,
    }
    if "team_id" not in values:
        values["team_id"] = values.get("team_code", certificate.team_code)
    fields = CERTIFICATE_TYPE_FIELDS.get(
        certificate.type.name,
        (
            "recipient_name",
            "event_name",
            "event_date",
            "event_venue",
            "team_id",
            "project_title",
            "domain_name",
        ),
    )
    custom_fonts = {
        ("Helvetica", False, False): "Helvetica",
        ("Helvetica", True, False): "Helvetica-Bold",
        ("Helvetica", False, True): "Helvetica-Oblique",
        ("Helvetica", True, True): "Helvetica-BoldOblique",
        ("Times", False, False): "Times-Roman",
        ("Times", True, False): "Times-Bold",
        ("Times", False, True): "Times-Italic",
        ("Times", True, True): "Times-BoldItalic",
        ("Courier", False, False): "Courier",
        ("Courier", True, False): "Courier-Bold",
        ("Courier", False, True): "Courier-Oblique",
        ("Courier", True, True): "Courier-BoldOblique",
    }
    for field in fields:
        if field in {"qr", "certificate_number"}:
            continue
        value = str(values.get(field, "")).strip()
        if not value:
            continue
        position = positions[field]
        font_name = custom_fonts[
            (position["font"], position["bold"], position["italic"])
        ]
        font_size = position["size"]
        text_x = width * position["x"] / 100
        max_width = max(1, 2 * min(text_x, width - text_x) - 4)
        while font_size > 1 and stringWidth(value, font_name, font_size) > max_width:
            font_size = max(1, font_size - 0.5)
        lines = simpleSplit(value, font_name, font_size, max_width)
        document.setFillColor(colors.HexColor(position["color"]))
        document.setFont(font_name, font_size)
        line_height = font_size * 1.5
        baseline = height * (1 - position["y"] / 100) - font_size * 0.35
        for line_number, line in enumerate(lines):
            line_offset = (line_number - (len(lines) - 1) / 2) * line_height
            draw_field_text = {
                "left": document.drawString,
                "center": document.drawCentredString,
                "right": document.drawRightString,
            }[position["align"]]
            draw_field_text(text_x, baseline - line_offset, line)

    for element in positions.get("custom_texts", []):
        text = re.sub(
            r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}",
            lambda match: str(values.get(match.group(1), "")),
            element["text"],
        ).strip()
        if not text:
            continue
        font_name = custom_fonts[
            (element["font"], element["bold"], element["italic"])
        ]
        font_size = element["size"]
        max_width = width * 0.9
        lines = simpleSplit(text, font_name, font_size, max_width)
        document.setFillColor(colors.HexColor(element["color"]))
        document.setFont(font_name, font_size)
        text_x = width * element["x"] / 100
        baseline = height * (1 - element["y"] / 100) - font_size * 0.35
        draw_line = {
            "left": document.drawString,
            "center": document.drawCentredString,
            "right": document.drawRightString,
        }[element["align"]]
        for line_number, line in enumerate(lines):
            line_offset = (line_number - (len(lines) - 1) / 2) * (
                font_size * 1.25
            )
            draw_line(text_x, baseline - line_offset, line)

    qr_image = qrcode.make(
        certificate_verification_url(certificate.certificate_id)
    )
    qr_buffer = io.BytesIO()
    qr_image.save(qr_buffer, format="PNG")
    qr_buffer.seek(0)
    qr = positions["qr"]
    qr_size = width * qr["size"] / 100
    qr_center_x = width * qr["x"] / 100
    qr_center_y = height * (1 - qr["y"] / 100)
    certificate_number = positions["certificate_number"]
    number_font = custom_fonts[
        (certificate_number["font"], certificate_number["bold"], certificate_number["italic"])
    ]
    number_size = min(certificate_number["size"], 7)
    number_max_width = max(
        1,
        min(
            width * 0.3,
            2 * min(qr_center_x, width - qr_center_x) - 4,
        ),
    )
    while (
        number_size > 4
        and stringWidth(certificate.certificate_id, number_font, number_size)
        > number_max_width
    ):
        number_size -= 0.25
    label_size = min(6, number_size)
    footer_height = label_size * 1.2 + number_size * 1.2 + 3
    qr_bottom = qr_center_y - qr_size / 2
    if qr_bottom < footer_height:
        qr_center_y = footer_height + qr_size / 2
        qr_bottom = footer_height
    document.drawImage(
        ImageReader(qr_buffer),
        qr_center_x - qr_size / 2,
        qr_bottom,
        width=qr_size,
        height=qr_size,
    )
    document.setFillColor(colors.HexColor(certificate_number["color"]))
    document.setFont(number_font, label_size)
    document.drawCentredString(
        qr_center_x,
        qr_bottom - 3 - label_size,
        "Certificate No.",
    )
    document.setFont(number_font, number_size)
    document.drawCentredString(
        qr_center_x,
        qr_bottom - 3 - label_size * 1.2 - number_size,
        certificate.certificate_id,
    )
    document.save()
    output.seek(0)
    return output


def certificate_files_from_uploads(files, selected_user_ids=None):
    members = {}
    for member in (
        TeamMember.query.join(User, TeamMember.user_id == User.id)
        .filter(User.active.is_(True))
        .all()
    ):
        if not member.user or (
            selected_user_ids is not None
            and member.user_id not in selected_user_ids
        ):
            continue
        for student_id in {member.prn, member.user.username}:
            normalized_id = student_id.strip().casefold()
            existing = members.get(normalized_id)
            if existing and existing.user_id != member.user_id:
                raise ValueError(
                    f"Student ID {student_id} is linked to more than one active account."
                )
            members[normalized_id] = member
    matches = []
    matched_student_ids = set()
    matched_user_ids = set()
    total_size = 0
    pdf_count = 0

    def add_pdf(filename, pdf_data):
        nonlocal total_size, pdf_count
        if len(pdf_data) < 8 or len(pdf_data) > MAX_CERTIFICATE_PDF_SIZE:
            raise ValueError(f"{filename} must be a PDF no larger than 25 MB.")
        total_size += len(pdf_data)
        pdf_count += 1
        if total_size > MAX_CERTIFICATE_ZIP_SIZE:
            raise ValueError("The combined certificate PDFs exceed 300 MB.")
        if pdf_count > MAX_CERTIFICATE_ZIP_ENTRIES:
            raise ValueError("You may upload no more than 500 certificate PDFs at once.")
        if not pdf_data.startswith(b"%PDF-"):
            raise ValueError(f"{filename} is not a valid PDF file.")
        try:
            with pymupdf.open(stream=pdf_data, filetype="pdf") as document:
                if document.needs_pass or document.page_count < 1:
                    raise ValueError(f"{filename} must be an unlocked PDF with at least one page.")
        except (RuntimeError, ValueError) as error:
            if isinstance(error, ValueError):
                raise
            raise ValueError(f"{filename} is not a readable PDF file.") from error

        student_id = filename[:-4].strip().casefold()
        member = members.get(student_id)
        if member is None:
            return
        if student_id in matched_student_ids or member.user_id in matched_user_ids:
            raise ValueError(
                f"More than one PDF matches the selected student {member.prn}."
            )
        matches.append((member, pdf_data))
        matched_student_ids.add(student_id)
        matched_user_ids.add(member.user_id)

    for uploaded_file in files:
        filename = os.path.basename(uploaded_file.filename or "").strip()
        if filename.lower().endswith(".zip"):
            uploaded_file.stream.seek(0)
            if not zipfile.is_zipfile(uploaded_file.stream):
                raise ValueError(f"{filename} is not a valid ZIP file.")
            uploaded_file.stream.seek(0)
            with zipfile.ZipFile(uploaded_file.stream) as archive:
                for info in archive.infolist():
                    if info.is_dir():
                        continue
                    archive_name = info.filename.replace("\\", "/")
                    member_filename = archive_name.rsplit("/", 1)[-1]
                    if not member_filename.lower().endswith(".pdf"):
                        continue
                    if stat.S_ISLNK(info.external_attr >> 16):
                        raise ValueError("The ZIP contains an invalid linked PDF entry.")
                    if info.file_size < 8 or info.file_size > MAX_CERTIFICATE_PDF_SIZE:
                        raise ValueError(
                            f"{member_filename} must be a PDF no larger than 25 MB."
                        )
                    if total_size + info.file_size > MAX_CERTIFICATE_ZIP_SIZE:
                        raise ValueError("The combined certificate PDFs exceed 300 MB.")
                    if pdf_count + 1 > MAX_CERTIFICATE_ZIP_ENTRIES:
                        raise ValueError("You may upload no more than 500 certificate PDFs at once.")
                    with archive.open(info) as pdf_file:
                        pdf_data = pdf_file.read(MAX_CERTIFICATE_PDF_SIZE + 1)
                    if len(pdf_data) != info.file_size:
                        raise ValueError(f"{member_filename} could not be read completely.")
                    add_pdf(member_filename, pdf_data)
        elif filename.lower().endswith(".pdf"):
            pdf_data = uploaded_file.stream.read(MAX_CERTIFICATE_PDF_SIZE + 1)
            add_pdf(filename, pdf_data)
        else:
            raise ValueError("Choose PDF files named Student_ID.pdf or one ZIP file.")
    if not matches:
        if selected_user_ids is not None:
            raise ValueError(
                "No uploaded PDF matched the selected active students. "
                "Check that each PDF filename is the student's exact Student ID."
            )
        raise ValueError(
            "No certificate PDFs matched an active student account. "
            "Name each file exactly Student_ID.pdf."
        )
    return matches


def add_certificate_verification_to_pdf(pdf_data, certificate_id):
    verification_url = certificate_verification_url(certificate_id)
    qr_image = qrcode.make(verification_url).convert("RGB")
    qr_buffer = io.BytesIO()
    qr_image.save(qr_buffer, format="PNG")

    try:
        document = pymupdf.open(stream=pdf_data, filetype="pdf")
    except (RuntimeError, ValueError) as error:
        raise ValueError("The certificate PDF could not be opened for QR stamping.") from error
    with document:
        if document.needs_pass or document.page_count < 1:
            raise ValueError("The certificate PDF must be unlocked and contain a page.")
        page = document[0]
        page_rect = page.rect
        margin = min(12, page_rect.width * 0.025, page_rect.height * 0.025)
        qr_size = min(78, page_rect.width * 0.14, page_rect.height * 0.15)
        panel_width = min(130, page_rect.width - 2 * margin)
        panel_height = qr_size + 30
        if panel_width <= 0 or panel_height + 2 * margin > page_rect.height:
            raise ValueError("The first PDF page is too small for a QR code and certificate number.")

        panel = pymupdf.Rect(
            page_rect.x0 + margin,
            page_rect.y1 - margin - panel_height,
            page_rect.x0 + margin + panel_width,
            page_rect.y1 - margin,
        )
        page.draw_rect(panel, color=None, fill=(1, 1, 1), overlay=True)
        qr_left = panel.x0 + (panel.width - qr_size) / 2
        qr_rect = pymupdf.Rect(qr_left, panel.y0, qr_left + qr_size, panel.y0 + qr_size)
        page.insert_image(qr_rect, stream=qr_buffer.getvalue(), overlay=True)

        label_rect = pymupdf.Rect(
            panel.x0 + 2,
            qr_rect.y1 + 1,
            panel.x1 - 2,
            qr_rect.y1 + 10,
        )
        page.insert_textbox(
            label_rect,
            "Certificate No.",
            fontname="helv",
            fontsize=6,
            align=pymupdf.TEXT_ALIGN_CENTER,
            color=(0, 0, 0),
            overlay=True,
        )
        number_rect = pymupdf.Rect(
            panel.x0 + 2,
            qr_rect.y1 + 10,
            panel.x1 - 2,
            panel.y1 - 1,
        )
        number_font_size = 6
        while (
            number_font_size > 3
            and stringWidth(
                certificate_id,
                "Helvetica",
                number_font_size,
            ) > number_rect.width
        ):
            number_font_size -= 0.25
        text_height = page.insert_textbox(
            number_rect,
            certificate_id,
            fontname="helv",
            fontsize=number_font_size,
            align=pymupdf.TEXT_ALIGN_CENTER,
            color=(0, 0, 0),
            overlay=True,
        )
        if text_height < 0:
            page.insert_textbox(
                number_rect,
                certificate_id,
                fontname="helv",
                fontsize=3,
                align=pymupdf.TEXT_ALIGN_CENTER,
                color=(0, 0, 0),
                overlay=True,
            )
        return document.tobytes(garbage=4, deflate=True)


def published_results():
    setting = db.session.get(SiteSetting, "nominations_published")
    return bool(setting and setting.value == "true") or (
        Nomination.query.filter_by(published=True).first() is not None
    )


def domain_result_published(domain_id):
    return public_leaderboard_published() or (
        DomainResultPublication.query.filter_by(domain_id=domain_id).first()
        is not None
    )


def ranked_domain_teams(domain_id):
    ranked = []
    teams = Team.query.filter_by(
        domain_id=domain_id, registration_status="Approved"
    ).all()
    for team in teams:
        assigned_count = JudgeAssignment.query.filter_by(team_id=team.id).count()
        finalized = Evaluation.query.filter_by(
            team_id=team.id, finalized=True, approved=True
        ).all()
        admin_evaluation = AdminEvaluation.query.filter_by(
            team_id=team.id, finalized=True
        ).first()
        if assigned_count and len(finalized) == assigned_count and admin_evaluation:
            score = sum(item.total for item in finalized) / assigned_count
            ranked.append((team, score + admin_evaluation.total))
    return sorted(ranked, key=lambda item: (-item[1], item[0].team_code))


def notify_team_stakeholders(team, title, body):
    recipient_ids = {
        user.id
        for user in User.query.filter(
            User.team_id == team.id,
            User.role.in_(("team_leader", "team_member")),
        ).all()
    }
    recipient_ids.update(
        assignment.mentor_id
        for assignment in MentorTeamAssignment.query.filter_by(team_id=team.id).all()
    )
    recipient_ids.update(
        assignment.coordinator_id
        for assignment in CoordinatorTeamAssignment.query.filter_by(
            team_id=team.id
        ).all()
    )
    db.session.add_all(
        UserNotification(user_id=user_id, title=title, body=body)
        for user_id in recipient_ids
    )


def notify_chat_message(recipient_ids, sender, scope):
    sender_name = sender.display_name or sender.username
    recipient_ids = set(recipient_ids) - {sender.id}
    db.session.add_all(
        UserNotification(
            user_id=user_id,
            title="New chat message",
            body=f"New message from {sender_name} · {scope}.",
        )
        for user_id in recipient_ids
    )


def public_leaderboard_published():
    setting = db.session.get(SiteSetting, "final_leaderboard_published")
    return bool(setting and setting.value == "true")


def eligible_certificate_recipients(certificate_type):
    normalized_name = certificate_type.name.casefold()
    if "participation" in normalized_name:
        return [
            (member.user, member.full_name, member.team.team_code)
            for member in TeamMember.query.join(User, TeamMember.user_id == User.id)
            .filter(User.active.is_(True)).all()
            if member.user
        ]
    if "mentor" in normalized_name:
        return [
            (user, user.display_name or user.username, "")
            for user in User.query.filter_by(role="mentor", active=True).all()
        ]
    if "coordinator" in normalized_name:
        return [
            (user, user.display_name or user.username, "")
            for user in User.query.filter_by(role="coordinator", active=True).all()
        ]
    if "excellence" in normalized_name or "winner" in normalized_name:
        winners = Winner.query.filter_by(published=True).all()
        winner_team_ids = {winner.team_id for winner in winners}
        return [
            (member.user, member.full_name, member.team.team_code)
            for member in TeamMember.query.filter(
                TeamMember.team_id.in_(winner_team_ids)
            ).join(User, TeamMember.user_id == User.id).all()
            if member.user
        ] if winner_team_ids else []
    return []


def certificate_visible_to_user(certificate, user):
    if user.role == "admin":
        return True
    if certificate.dispatch_status != "dispatched":
        return False
    if certificate.user_id:
        return certificate.user_id == user.id
    if user.role == "team_leader" and user.team_id:
        team = db.session.get(Team, user.team_id)
        return bool(
            team
            and certificate.team_code == team.team_code
            and certificate.recipient_name == (team.leader_name or team.name)
        )
    if user.role == "team_member" and user.member_record:
        return (
            certificate.team_code == user.member_record.team.team_code
            and certificate.recipient_name == user.member_record.full_name
        )
    return False


def certificate_download_filename(certificate, extension):
    student_id = ""
    if certificate.user_id:
        member = TeamMember.query.filter_by(user_id=certificate.user_id).first()
        if member:
            student_id = member.prn
    if not student_id:
        try:
            fields = json.loads(certificate.fields_json or "{}")
        except (TypeError, json.JSONDecodeError):
            fields = {}
        if isinstance(fields, dict):
            student_id = fields.get("student_id", "")
    filename = secure_filename(str(student_id).strip()) or certificate.certificate_id
    return f"{filename}.{extension}"


def certificates_for_user(user):
    filters = [
        and_(
            Certificate.user_id == user.id,
            Certificate.dispatch_status == "dispatched",
        ),
    ]
    team = db.session.get(Team, user.team_id) if user.team_id else None
    recipient_name = None
    if user.role == "team_leader" and team:
        recipient_name = team.leader_name or team.name
    elif user.role == "team_member" and user.member_record:
        team = user.member_record.team
        recipient_name = user.member_record.full_name
    if team and recipient_name:
        filters.append(and_(
            Certificate.user_id.is_(None),
            Certificate.team_code == team.team_code,
            Certificate.recipient_name == recipient_name,
            Certificate.dispatch_status == "dispatched",
        ))
    return Certificate.query.filter(or_(*filters)).order_by(
        Certificate.issued_at.desc(), Certificate.id.desc()
    ).all()


@main.route("/")
def home():
    domains = Domain.query.order_by(Domain.id).all()
    teams = Team.query.count()
    announcements = Announcement.query.order_by(Announcement.created_at.desc()).limit(5).all()
    return render_template("public/home.html", domains=domains, teams=teams, announcements=announcements)


@main.route("/healthz")
def health_check():
    try:
        db.session.execute(select(1))
    except SQLAlchemyError:
        db.session.rollback()
        current_app.logger.exception("Application database health check failed.")
        return jsonify({"status": "unavailable"}), 503
    return jsonify({"status": "ok"})


@main.route("/domains/<int:domain_id>")
def public_domain_details(domain_id):
    domain = db.get_or_404(Domain, domain_id)
    teams = Team.query.filter_by(domain_id=domain.id).order_by(Team.team_code).all()
    is_admin = current_user.is_authenticated and current_user.role == "admin"
    domain_coordinators = CoordinatorDomainAssignment.query.filter_by(
        domain_id=domain.id
    ).all() if is_admin else []
    team_details = []
    for team in teams:
        detail = {"team": team}
        if is_admin:
            mentor_assignments = MentorTeamAssignment.query.filter_by(
                team_id=team.id
            ).all()
            coordinator_assignments = CoordinatorTeamAssignment.query.filter_by(
                team_id=team.id
            ).all()
            assigned_coordinator_ids = {
                item.coordinator_id for item in coordinator_assignments
            }
            domain_coordinator_users = [
                item.coordinator
                for item in domain_coordinators
                if item.coordinator_id not in assigned_coordinator_ids
            ]
            judge_assignments = JudgeAssignment.query.filter_by(
                team_id=team.id
            ).all()
            detail.update(
                mentor_assignments=mentor_assignments,
                coordinator_assignments=coordinator_assignments,
                domain_coordinator_users=domain_coordinator_users,
                judge_assignments=judge_assignments,
                mentor_names=[
                    item.mentor.display_name or item.mentor.username
                    for item in mentor_assignments
                ],
                coordinator_names=[
                    item.coordinator.display_name or item.coordinator.username
                    for item in coordinator_assignments
                ] + [
                    user.display_name or user.username
                    for user in domain_coordinator_users
                ],
                judge_names=[
                    item.judge.display_name or item.judge.username
                    for item in judge_assignments
                ],
                evaluations=Evaluation.query.filter_by(
                    team_id=team.id
                ).order_by(Evaluation.submitted_at.desc()).all(),
                admin_score=AdminEvaluation.query.filter_by(
                    team_id=team.id
                ).first(),
                mentor_scores=MentorEvaluation.query.filter_by(
                    team_id=team.id
                ).order_by(MentorEvaluation.submitted_at.desc()).all(),
                leader_account=User.query.filter_by(
                    team_id=team.id, role="team_leader"
                ).first(),
            )
        team_details.append(detail)
    return render_template(
        "public/domain.html",
        domain=domain,
        team_details=team_details,
        is_admin=is_admin,
    )


@main.route("/announcements")
def public_announcements():
    return render_template(
        "public/announcements.html",
        announcements=Announcement.query.order_by(Announcement.created_at.desc()).all(),
    )


@main.route("/rankings")
def rankings():
    data = []
    published = public_leaderboard_published()
    admin_preview = current_user.is_authenticated and current_user.role == "admin"
    for domain in Domain.query.order_by(Domain.id).all():
        released = domain_result_published(domain.id)
        rows = (
            ranked_domain_teams(domain.id)
            if released or admin_preview
            else []
        )
        data.append((domain, rows, released))
    return render_template(
        "public/rankings.html",
        data=data,
        published=published,
        admin_preview=admin_preview,
        visible_domain_results=any(released for _, _, released in data),
    )

@main.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("main.dashboard"))
    if request.method == "POST":
        ip_address = request.remote_addr or "unknown"
        cutoff = datetime.utcnow() - LOGIN_WINDOW
        LoginAttempt.query.filter(LoginAttempt.created_at < cutoff).delete()
        recent_failures = LoginAttempt.query.filter_by(
            ip_address=ip_address, successful=False
        ).filter(LoginAttempt.created_at >= cutoff).count()
        if recent_failures >= MAX_LOGIN_ATTEMPTS:
            db.session.commit()
            flash("Too many login attempts. Please try again in 15 minutes.", "danger")
            return render_template("auth/login.html"), 429

        username = request.form.get("username", "").strip()
        selected_role = request.form.get("role", "")
        user = User.query.filter_by(username=username).first()
        if (
            selected_role in USER_ROLES
            and user
            and user.active
            and user.role == selected_role
            and user.check_password(request.form.get("password", ""))
        ):
            LoginAttempt.query.filter_by(ip_address=ip_address, successful=False).delete()
            db.session.add(LoginAttempt(ip_address=ip_address, successful=True))
            db.session.add(AuditLog(
                user_id=user.id,
                action="Logged in",
                details="Successful login",
            ))
            db.session.commit()
            session.clear()
            login_user(user)
            return redirect(url_for("main.dashboard"))
        db.session.add(LoginAttempt(ip_address=ip_address, successful=False))
        db.session.commit()
        flash("Invalid role, username/password, or inactive account.", "danger")
    return render_template("auth/login.html", roles=sorted(USER_ROLES))


@main.route("/logout", methods=["POST"])
@login_required
def logout():
    admin_id = session.get("admin_impersonation_admin_id")
    if admin_id:
        db.session.add(AuditLog(
            user_id=admin_id,
            action="Ended impersonation by logout",
            details=f"Target user ID {current_user.id} ({current_user.username})",
        ))
        for key in (
            "admin_impersonation_admin_id",
            "admin_impersonation_admin_session_id",
            "admin_impersonation_target_id",
            "admin_impersonation_mode",
        ):
            session.pop(key, None)
    else:
        log("Logged out")
    db.session.commit()
    logout_user()
    return redirect(url_for("main.home"))


@main.route(
    "/admin/users/<int:user_id>/impersonate",
    methods=["POST"],
)
@roles_required("admin")
def start_admin_impersonation(user_id):
    if session.get("admin_impersonation_admin_id"):
        abort(409, description="End the current impersonation before selecting another account.")
    target = db.get_or_404(User, user_id)
    if (
        not target.active
        or target.id == current_user.id
        or target.role == "admin"
    ):
        abort(404)
    mode = request.form.get("mode", "")
    if mode not in {"view_only", "full_access"}:
        abort(400, description="Choose view-only or full audited access.")

    administrator = current_user._get_current_object()
    db.session.add(AuditLog(
        user_id=administrator.id,
        action="Started admin impersonation",
        details=(
            f"Target user ID {target.id} ({target.username}); "
            f"mode={mode}; no fixed timeout"
        ),
    ))
    db.session.commit()

    session["admin_impersonation_admin_id"] = administrator.id
    session["admin_impersonation_admin_session_id"] = administrator.get_id()
    session["admin_impersonation_target_id"] = target.id
    session["admin_impersonation_mode"] = mode
    logout_user()
    login_user(target, remember=False, fresh=True)
    flash(
        f"Now viewing {target.display_name or target.username} "
        f"({mode.replace('_', ' ')}).",
        "warning",
    )
    return redirect(url_for("main.dashboard"))


@main.route("/admin/impersonation/upgrade", methods=["POST"])
@login_required
def upgrade_admin_impersonation():
    admin_id = session.get("admin_impersonation_admin_id")
    target_id = session.get("admin_impersonation_target_id")
    admin_session_id = session.get("admin_impersonation_admin_session_id")
    if (
        not admin_id
        or target_id != current_user.id
        or session.get("admin_impersonation_mode") != "view_only"
        or not admin_session_id
    ):
        abort(404)
    administrator = db.session.get(User, admin_id)
    if (
        not administrator
        or not administrator.active
        or administrator.role != "admin"
        or administrator.get_id() != admin_session_id
    ):
        logout_user()
        session.clear()
        flash("The original admin session is no longer valid. Please sign in again.", "warning")
        return redirect(url_for("main.login"))

    db.session.add(AuditLog(
        user_id=administrator.id,
        action="Upgraded admin impersonation to full access",
        details=f"Target user ID {current_user.id} ({current_user.username})",
    ))
    db.session.commit()
    session["admin_impersonation_mode"] = "full_access"
    flash(
        f"Full audited access is now enabled for {current_user.display_name or current_user.username}.",
        "warning",
    )
    return redirect(url_for("main.dashboard"))


@main.route("/admin/impersonation/end", methods=["POST"])
@login_required
def end_admin_impersonation():
    admin_id = session.get("admin_impersonation_admin_id")
    target_id = session.get("admin_impersonation_target_id")
    admin_session_id = session.get("admin_impersonation_admin_session_id")
    if (
        not admin_id
        or target_id != current_user.id
        or not admin_session_id
    ):
        abort(404)
    administrator = db.session.get(User, admin_id)
    if (
        not administrator
        or not administrator.active
        or administrator.role != "admin"
        or administrator.get_id() != admin_session_id
    ):
        logout_user()
        session.clear()
        flash("The original admin session is no longer valid. Please sign in again.", "warning")
        return redirect(url_for("main.login"))

    db.session.add(AuditLog(
        user_id=administrator.id,
        action="Ended admin impersonation",
        details=f"Target user ID {current_user.id} ({current_user.username})",
    ))
    db.session.commit()
    for key in (
        "admin_impersonation_admin_id",
        "admin_impersonation_admin_session_id",
        "admin_impersonation_target_id",
        "admin_impersonation_mode",
    ):
        session.pop(key, None)
    logout_user()
    login_user(administrator, remember=False, fresh=True)
    flash("Returned to the administrator account.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/account/profile", methods=["GET", "POST"])
@login_required
def account_profile():
    member = current_user.member_record if current_user.role == "team_member" else None
    profile_teams = (
        [db.session.get(Team, current_user.team_id)]
        if current_user.team_id
        else assigned_teams_for_user(current_user)
    )
    profile_teams = [team for team in profile_teams if team]
    if request.method == "POST":
        display_name = request.form.get("display_name", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        if not display_name or len(display_name) > 150:
            flash("Enter a display name of at most 150 characters.", "danger")
            return render_template(
                "auth/profile.html", member=member, profile_teams=profile_teams
            ), 400
        if len(email) > 150 or (
            email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)
        ):
            flash("Enter a valid email address of at most 150 characters.", "danger")
            return render_template(
                "auth/profile.html", member=member, profile_teams=profile_teams
            ), 400
        if len(phone) > 40 or (
            phone and not re.fullmatch(r"[0-9+().\-\s]+", phone)
        ):
            flash("Enter a valid phone number of at most 40 characters.", "danger")
            return render_template(
                "auth/profile.html", member=member, profile_teams=profile_teams
            ), 400

        current_user.display_name = display_name
        current_user.email = email
        current_user.phone = phone
        if member:
            member.full_name = display_name
            member.email = email
            member.phone = phone
            for field, limit in (("roll_number", 40), ("department", 100), ("year", 30)):
                value = request.form.get(field, "").strip()
                if len(value) > limit:
                    flash(
                        f"{field.replace('_', ' ').title()} must be at most {limit} characters.",
                        "danger",
                    )
                    return render_template(
                        "auth/profile.html",
                        member=member,
                        profile_teams=profile_teams,
                    ), 400
                setattr(member, field, value)
        log("Updated own account profile")
        db.session.commit()
        flash("Your profile has been updated.", "success")
        return redirect(url_for("main.account_profile"))
    return render_template(
        "auth/profile.html", member=member, profile_teams=profile_teams
    )


@main.route("/account/password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        if not current_user.check_password(request.form.get("current_password", "")):
            flash("Current password is incorrect.", "danger")
            return render_template("auth/change_password.html"), 400
        password = request.form.get("new_password", "")
        confirmation = request.form.get("confirm_password", "")
        if len(password) < 4:
            flash("Passwords must be at least 4 characters long.", "danger")
        elif password != confirmation:
            flash("New password and confirmation do not match.", "danger")
        else:
            current_user.set_password(password)
            log("Changed password")
            db.session.commit()
            login_user(current_user)
            flash("Password changed successfully.", "success")
            return redirect(url_for("main.dashboard"))
    return render_template("auth/change_password.html")


@main.route("/account/history")
@login_required
def account_history():
    selected_user = current_user
    requested_user_id = request.args.get("user_id", "").strip()
    page = request.args.get("page", 1, type=int)
    if page is None or page < 1:
        abort(400, description="Choose a valid history page.")
    if current_user.role == "admin":
        if requested_user_id:
            try:
                selected_user_id = int(requested_user_id)
            except ValueError:
                abort(400, description="Choose a valid account.")
            selected_user = db.get_or_404(User, selected_user_id)
        else:
            selected_user = None
        users = User.query.order_by(User.role, User.display_name, User.username).all()
    else:
        if requested_user_id and requested_user_id != str(current_user.id):
            abort(403)
        selected_user_id = current_user.id
        users = []

    audit_select = select(
        AuditLog.created_at.label("created_at"),
        literal("Activity").label("kind"),
        AuditLog.action.label("title"),
        AuditLog.details.label("details"),
        AuditLog.user_id.label("user_id"),
        AuditLog.id.label("entry_id"),
    )
    notification_select = select(
        UserNotification.created_at.label("created_at"),
        literal("Notification").label("kind"),
        UserNotification.title.label("title"),
        UserNotification.body.label("details"),
        UserNotification.user_id.label("user_id"),
        UserNotification.id.label("entry_id"),
    )
    if current_user.role != "admin" or requested_user_id:
        audit_select = audit_select.where(AuditLog.user_id == selected_user_id)
        notification_select = notification_select.where(
            UserNotification.user_id == selected_user_id
        )
    history_query = union_all(audit_select, notification_select).subquery()
    total_history = db.session.scalar(
        select(db.func.count()).select_from(history_query)
    ) or 0
    page_size = 50
    history_rows = db.session.execute(
        select(history_query)
        .order_by(
            history_query.c.created_at.desc(),
            history_query.c.kind,
            history_query.c.entry_id.desc(),
        )
        .offset((page - 1) * page_size)
        .limit(page_size)
    ).mappings().all()
    audit_users = {
        user.id: user
        for user in User.query.filter(
            User.id.in_({
                row["user_id"] for row in history_rows
                if row["user_id"] is not None
            })
        ).all()
    } if any(row["user_id"] is not None for row in history_rows) else {}
    history = [
        {
            "created_at": row["created_at"],
            "kind": row["kind"],
            "title": row["title"],
            "details": row["details"] or "",
            "username": (
                (
                    audit_users[row["user_id"]].display_name
                    or audit_users[row["user_id"]].username
                )
                if row["user_id"] in audit_users
                else "Account no longer available"
            ),
        }
        for row in history_rows
    ]
    return render_template(
        "account/history.html",
        history=history,
        users=users,
        selected_user=selected_user,
        administrator=current_user.role == "admin",
        page=page,
        page_count=max(1, math.ceil(total_history / page_size)),
    )


@main.route("/account/notifications/<int:notification_id>/read", methods=["POST"])
@login_required
def mark_notification_read(notification_id):
    notification = UserNotification.query.filter_by(
        id=notification_id,
        user_id=current_user.id,
    ).first_or_404()
    notification.read = True
    db.session.commit()
    return redirect(request.referrer or url_for("main.account_history"))


@main.route("/dashboard")
@login_required
def dashboard():
    role = current_user.role
    if role == "admin":
        return redirect(url_for("main.admin_dashboard"))
    if role == "team_leader":
        return redirect(url_for("main.team_dashboard"))
    if role == "team_member":
        return redirect(url_for("main.member_dashboard"))
    if role == "mentor":
        return redirect(url_for("main.simple_dashboard", role="Mentor"))
    if role == "coordinator":
        return redirect(url_for("main.simple_dashboard", role="Coordinator"))
    if role == "judge":
        return redirect(url_for("main.judge_dashboard"))
    abort(403)

@main.route("/admin")
@roles_required("admin")
def admin_dashboard():
    all_accounts = User.query.order_by(
        User.role, User.display_name, User.username
    ).all()
    all_teams = Team.query.order_by(Team.team_code).all()
    team_labels = {team.id: team.team_code for team in all_teams}
    account_team_labels = {
        account.id: [team_labels[account.team_id]]
        if account.team_id in team_labels
        else []
        for account in all_accounts
    }
    for assignment in MentorTeamAssignment.query.all():
        account_team_labels.setdefault(
            assignment.mentor_id, []
        ).append(assignment.team.team_code)
    for assignment in CoordinatorTeamAssignment.query.all():
        account_team_labels.setdefault(
            assignment.coordinator_id, []
        ).append(assignment.team.team_code)
    account_contacts = {
        account.id: {
            "email": user_notification_email(account),
            "phone": user_notification_phone(account),
        }
        for account in all_accounts
    }
    ranked_domains = [
        (domain, ranked_domain_teams(domain.id))
        for domain in Domain.query.order_by(Domain.id).all()
    ]
    evaluation_progress = []
    for team in all_teams:
        assigned_judges = JudgeAssignment.query.filter_by(team_id=team.id).count()
        finalized_judges = Evaluation.query.filter_by(
            team_id=team.id, finalized=True, approved=True
        ).count()
        submitted_judges = Evaluation.query.filter_by(
            team_id=team.id, finalized=True
        ).count()
        admin_score = AdminEvaluation.query.filter_by(team_id=team.id).first()
        latest_mentor_proposal = MentorEvaluation.query.filter_by(
            team_id=team.id
        ).order_by(MentorEvaluation.submitted_at.desc()).first()
        evaluation_progress.append({
            "team": team,
            "assigned_judges": assigned_judges,
            "finalized_judges": finalized_judges,
            "submitted_judges": submitted_judges,
            "admin_score": admin_score,
            "mentor_proposal": latest_mentor_proposal,
        })
    nominations_generated = db.session.get(SiteSetting, "nominations_generated")
    return render_template("admin/dashboard.html",
        teams=Team.query.count(), users=User.query.count(),
        members=TeamMember.query.count(), videos=ProjectVideo.query.count(),
        pending=Team.query.filter_by(registration_status="Pending").count(),
        evaluations=Evaluation.query.filter_by(finalized=True).count(),
        pending_judge_evaluations=Evaluation.query.filter_by(
            finalized=True, approved=False
        ).count(),
        judge_evaluations_pending_review=Evaluation.query.filter_by(
            finalized=True, approved=False
        ).order_by(Evaluation.submitted_at.asc()).all(),
        admin_evaluations=AdminEvaluation.query.count(),
        pending_mentor_evaluations=MentorEvaluation.query.filter_by(status="Pending").count(),
        mentor_evaluations=MentorEvaluation.query.order_by(
            MentorEvaluation.submitted_at.desc()
        ).all(),
        nominations=Nomination.query.count(), winners=Winner.query.count(),
        domains=Domain.query.order_by(Domain.id).all(),
        all_teams=all_teams,
        evaluation_progress=evaluation_progress,
        admin_score_by_team={
            entry["team"].id: entry["admin_score"]
            for entry in evaluation_progress
        },
        ranked_domains=ranked_domains,
        all_accounts=all_accounts,
        account_team_labels=account_team_labels,
        account_contacts=account_contacts,
        mentor_count=sum(1 for user in all_accounts if user.role == "mentor"),
        coordinator_count=sum(
            1 for user in all_accounts if user.role == "coordinator"
        ),
        nominations_published=published_results(),
        nominations_generated=bool(
            nominations_generated and nominations_generated.value == "true"
        ),
        domain_publications={
            publication.domain_id: publication
            for publication in DomainResultPublication.query.all()
        },
        final_leaderboard_published=public_leaderboard_published())


@main.route("/admin/overview/<section>")
@roles_required("admin")
def admin_overview(section):
    section_titles = {
        "mentors": "Mentors",
        "coordinators": "Coordinators",
        "teams": "Teams",
        "members": "Members",
        "pending-teams": "Pending teams",
        "judge-pending": "Judge marks awaiting approval",
        "judge-approved": "Judge evaluation marks & approval",
        "admin-scores": "Admin 40-mark entries",
        "mentor-proposals": "Mentor 40-mark proposals",
        "videos": "Project videos",
        "winners": "Winners",
        "scores": "100-mark evaluation",
    }
    if section not in section_titles:
        abort(404)
    if section in {"scores", "admin-scores", "judge-pending", "judge-approved", "mentor-proposals"}:
        teams = Team.query.order_by(Team.team_code).all()
        score_rows = []
        for team in teams:
            admin_score = AdminEvaluation.query.filter_by(team_id=team.id).first()
            mentor_proposals = MentorEvaluation.query.filter_by(
                team_id=team.id
            ).order_by(MentorEvaluation.submitted_at.desc()).all()
            judge_evaluations = Evaluation.query.filter_by(
                team_id=team.id
            ).order_by(Evaluation.submitted_at.desc()).all()
            if section == "judge-approved":
                judge_evaluations = [
                    item for item in judge_evaluations if item.finalized
                ]
            forty_score = admin_score.total if admin_score else (
                mentor_proposals[0].total if mentor_proposals else None
            )
            judge_scores = [evaluation.total for evaluation in judge_evaluations]
            sixty_score = (
                sum(judge_scores) / len(judge_scores) if judge_scores else None
            )
            total_score = (
                forty_score + sixty_score
                if forty_score is not None and sixty_score is not None
                else None
            )
            if section == "admin-scores" and not admin_score:
                continue
            if section == "mentor-proposals" and not mentor_proposals:
                continue
            if section == "judge-pending" and not any(
                item.finalized and not item.approved for item in judge_evaluations
            ):
                continue
            if section == "judge-approved" and not judge_evaluations:
                continue
            score_rows.append({
                "team": team,
                "admin_score": admin_score,
                "mentor_proposals": mentor_proposals,
                "judge_evaluations": judge_evaluations,
                "forty_score": forty_score,
                "sixty_score": sixty_score,
                "total_score": total_score,
            })
        return render_template(
            "admin/overview.html",
            section=section,
            title=section_titles[section],
            score_rows=score_rows,
            nominations_published=published_results(),
        )

    users = []
    teams = []
    members = []
    videos = []
    winner_records = []
    if section in {"mentors", "coordinators"}:
        role = "mentor" if section == "mentors" else "coordinator"
        users = User.query.filter_by(role=role).order_by(
            User.display_name, User.username
        ).all()
    elif section in {"teams", "pending-teams"}:
        query = Team.query.order_by(Team.team_code)
        if section == "pending-teams":
            query = query.filter_by(registration_status="Pending")
        teams = query.all()
    elif section == "members":
        members = TeamMember.query.join(Team).order_by(
            Team.team_code, TeamMember.full_name
        ).all()
    elif section == "videos":
        videos = ProjectVideo.query.order_by(ProjectVideo.uploaded_at.desc()).all()
    elif section == "winners":
        winner_records = Winner.query.join(Team).order_by(Team.team_code).all()
    return render_template(
        "admin/overview.html",
        section=section,
        title=section_titles[section],
        users=users,
        teams=teams,
        members=members,
        videos=videos,
        winners=winner_records,
    )


@main.route("/admin/data")
@roles_required("admin")
def admin_data():
    return render_template(
        "admin/data.html",
        domains=Domain.query.order_by(Domain.id).all(),
    )


@main.route("/admin/data/backup")
@roles_required("admin")
def download_event_backup():
    archive_buffer = io.BytesIO()
    dialect = db.engine.dialect.name
    table_names = [table.name for table in db.metadata.sorted_tables]
    included_uploads = []
    with zipfile.ZipFile(
        archive_buffer, "w", compression=zipfile.ZIP_DEFLATED
    ) as archive:
        if dialect == "sqlite":
            with tempfile.TemporaryDirectory() as backup_directory:
                snapshot_path = os.path.join(
                    backup_directory, "future_tech_innovision.sqlite"
                )
                source_connection = db.engine.raw_connection()
                destination_connection = sqlite3.connect(snapshot_path)
                try:
                    source_connection.backup(destination_connection)
                finally:
                    destination_connection.close()
                    source_connection.close()
                archive.write(
                    snapshot_path,
                    "database/future_tech_innovision.sqlite",
                )

        schema_sql = "\n\n".join(
            str(CreateTable(table).compile(dialect=db.engine.dialect)) + ";"
            for table in db.metadata.sorted_tables
        )
        archive.writestr("database/schema.sql", schema_sql)

        with db.engine.connect() as connection:
            for table in db.metadata.sorted_tables:
                csv_path = f"data/{table.name}.csv"
                with archive.open(csv_path, "w") as csv_bytes:
                    text_stream = io.TextIOWrapper(
                        csv_bytes, encoding="utf-8-sig", newline=""
                    )
                    writer = csv.writer(text_stream)
                    column_names = [column.name for column in table.columns]
                    writer.writerow(column_names)
                    result = connection.execution_options(
                        stream_results=True
                    ).execute(table.select())
                    for row in result:
                        writer.writerow([
                            ""
                            if table.name == "user"
                            and column_name == "password_hash"
                            else (
                                ""
                                if row._mapping[column_name] is None
                                else str(row._mapping[column_name])
                            )
                            for column_name in column_names
                        ])
                    text_stream.flush()
                    text_stream.detach()

        upload_folder = current_app.config["UPLOAD_FOLDER"]
        if os.path.isdir(upload_folder):
            for root, directory_names, filenames in os.walk(
                upload_folder, followlinks=False
            ):
                directory_names[:] = [
                    name for name in directory_names
                    if not os.path.islink(os.path.join(root, name))
                ]
                for filename in filenames:
                    file_path = os.path.join(root, filename)
                    if os.path.islink(file_path) or not os.path.isfile(file_path):
                        continue
                    relative_path = os.path.relpath(file_path, upload_folder)
                    archive_path = os.path.join("uploads", relative_path).replace(
                        os.sep, "/"
                    )
                    archive.write(file_path, archive_path)
                    included_uploads.append(archive_path)

        archive.writestr(
            "manifest.json",
            json.dumps(
                {
                    "created_at": datetime.utcnow().isoformat() + "Z",
                    "database_dialect": dialect,
                    "database_snapshot_included": dialect == "sqlite",
                    "table_exports": table_names,
                    "password_hashes_in_user_csv": "redacted",
                    "uploaded_files": included_uploads,
                },
                indent=2,
            ),
        )

    archive_buffer.seek(0)
    return send_file(
        archive_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name="future-tech-innovision-2026-backup.zip",
    )


@main.route("/admin/monitor")
@roles_required("admin")
def admin_monitor():
    return render_template(
        "admin/monitor.html",
        teams=Team.query.order_by(Team.team_code).all(),
    )

@main.route("/admin/teams", methods=["GET", "POST"])
@roles_required("admin")
def admin_teams():
    if request.method == "POST":
        if published_results():
            flash("Team details are frozen after nominations are published.", "warning")
            return redirect(url_for("main.admin_teams"))
        team = Team.query.filter_by(team_code=request.form["team_code"]).first_or_404()
        team.name = request.form.get("name", team.team_code)
        team.leader_name = request.form.get("leader_name", "")
        team.email = request.form.get("email", "")
        team.phone = request.form.get("phone", "")
        team.department = request.form.get("department", "")
        team.year = request.form.get("year", "")
        domain_id = request.form.get("domain_id", "")
        if domain_id:
            try:
                domain_id = int(domain_id)
            except ValueError:
                abort(400, description="Invalid technical domain.")
            if not db.session.get(Domain, domain_id):
                abort(400, description="Invalid technical domain.")
        else:
            domain_id = None
        team.domain_id = domain_id
        team.project_title = request.form.get("project_title", "")
        team.abstract = str(sanitize_rich_text(request.form.get("abstract", "")))
        team.problem_statement = str(sanitize_rich_text(request.form.get("problem_statement", "")))
        team.objectives = str(sanitize_rich_text(request.form.get("objectives", "")))
        team.technologies = str(sanitize_rich_text(request.form.get("technologies", "")))
        status = request.form.get("registration_status", "Pending")
        if status not in REGISTRATION_STATUSES:
            abort(400, description="Invalid registration status.")
        team.registration_status = status
        log("Updated team", team.team_code)
        db.session.commit()
        flash("Team updated.", "success")
        return redirect(url_for("main.admin_teams"))
    return render_template(
        "admin/teams.html",
        teams=Team.query.order_by(Team.team_code).all(),
        domains=Domain.query.all(),
        admin_scores={
            evaluation.team_id: evaluation
            for evaluation in AdminEvaluation.query.all()
        },
    )


@main.route("/admin/teams/create", methods=["POST"])
@roles_required("admin")
def create_team():
    if published_results():
        abort(409, description="Teams cannot be added after nominations are published.")
    team_code = request.form.get("team_code", "").strip().upper()
    name = request.form.get("name", "").strip()
    leader_name = request.form.get("leader_name", "").strip()
    domain_id_value = request.form.get("domain_id", "").strip()
    if not re.fullmatch(r"FT_[0-9]{2,3}", team_code):
        abort(400, description="Team IDs must use the format FT_28.")
    if len(name) > 150 or len(leader_name) > 150:
        abort(400, description="Team and leader names must be at most 150 characters.")
    if Team.query.filter_by(team_code=team_code).first():
        flash(f"Team {team_code} already exists.", "danger")
        return redirect(url_for("main.admin_teams"))
    if User.query.filter_by(username=team_code).first():
        flash(f"An account named {team_code} already exists.", "danger")
        return redirect(url_for("main.admin_teams"))
    domain = None
    if domain_id_value:
        try:
            domain = db.session.get(Domain, int(domain_id_value))
        except ValueError:
            abort(400, description="Invalid technical domain.")
        if domain is None:
            abort(400, description="Invalid technical domain.")
    initial_password = secrets.token_urlsafe(9)[:12]
    team = Team(
        team_code=team_code,
        name=name or team_code,
        leader_name=leader_name,
        domain=domain,
    )
    db.session.add(team)
    db.session.flush()
    leader = User(
        username=team_code,
        role="team_leader",
        display_name=leader_name or name or team_code,
        team_id=team.id,
    )
    leader.set_password(initial_password)
    db.session.add(leader)
    log("Created team and leader account", team_code)
    db.session.commit()
    flash(
        f"Created {team_code}. Team leader login: {team_code}; "
        f"temporary password: {initial_password}. Save this password now.",
        "success",
    )
    return redirect(url_for("main.admin_teams"))


@main.route("/admin/teams/<int:team_id>/delete", methods=["POST"])
@roles_required("admin")
def delete_team(team_id):
    team = db.get_or_404(Team, team_id)
    if published_results() or DomainResultPublication.query.filter_by(
        domain_id=team.domain_id
    ).first():
        abort(409, description="Teams are frozen after nominations or domain results are published.")
    if (
        Evaluation.query.filter_by(team_id=team.id).first()
        or AdminEvaluation.query.filter_by(team_id=team.id).first()
        or MentorEvaluation.query.filter_by(team_id=team.id).first()
        or Winner.query.filter_by(team_id=team.id).first()
        or Attendance.query.join(TeamMember).filter(
            TeamMember.team_id == team.id
        ).first()
    ):
        flash(
            "This team has evaluation or attendance history. Keep it and change its registration status instead.",
            "danger",
        )
        return redirect(url_for("main.admin_teams"))

    video_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "videos")
    video_filenames = [
        video.filename for video in ProjectVideo.query.filter_by(team_id=team.id).all()
    ]
    submission_filenames = [
        submission.filename
        for submission in TeamSubmission.query.filter_by(team_id=team.id).all()
    ]
    member_records = TeamMember.query.filter_by(team_id=team.id).all()
    user_ids = {
        user.id
        for user in User.query.filter_by(team_id=team.id).all()
        if user.role in {"team_leader", "team_member"}
    }
    user_ids.update(
        member.user_id for member in member_records if member.user_id is not None
    )
    for user_id in user_ids:
        AuditLog.query.filter_by(user_id=user_id).update(
            {"user_id": None}, synchronize_session=False
        )
        UserNotification.query.filter_by(user_id=user_id).delete()
    for member in member_records:
        if member.user_id in user_ids:
            member.user_id = None
        db.session.delete(member)
    User.query.filter(User.id.in_(user_ids)).delete(synchronize_session=False)
    User.query.filter_by(team_id=team.id).update(
        {"team_id": None}, synchronize_session=False
    )
    JudgeAssignment.query.filter_by(team_id=team.id).delete()
    MentorTeamAssignment.query.filter_by(team_id=team.id).delete()
    CoordinatorTeamAssignment.query.filter_by(team_id=team.id).delete()
    ProjectVideo.query.filter_by(team_id=team.id).delete()
    TeamSubmission.query.filter_by(team_id=team.id).delete()
    Nomination.query.filter_by(team_id=team.id).delete()
    team_code = team.team_code
    db.session.delete(team)
    log("Deleted team", team_code)
    db.session.commit()
    for filename in video_filenames:
        if os.path.basename(filename) == filename:
            path = os.path.join(video_directory, filename)
            if os.path.isfile(path):
                os.remove(path)
    submission_directory = os.path.join(
        current_app.config["UPLOAD_FOLDER"], "submissions"
    )
    for filename in submission_filenames:
        if os.path.basename(filename) == filename:
            path = os.path.join(submission_directory, filename)
            if os.path.isfile(path):
                os.remove(path)
    flash(f"Team {team_code} and its team accounts were deleted.", "success")
    return redirect(url_for("main.admin_teams"))

@main.route("/admin/users", methods=["GET", "POST"])
@roles_required("admin")
def admin_users():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        role = request.form.get("role", "")
        password = request.form.get("password", "")
        display_name = request.form.get("display_name", "").strip()
        email = request.form.get("email", "").strip()
        phone = request.form.get("phone", "").strip()
        existing_user = User.query.filter_by(username=username).first()
        if not re.fullmatch(r"[A-Za-z0-9_.-]{3,80}", username):
            flash("Username must be 3–80 letters, numbers, dots, underscores or hyphens.", "danger")
        elif role not in USER_ROLES:
            abort(400, description="Invalid account role.")
        elif len(email) > 150 or (
            email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)
        ):
            flash("Enter a valid email address of at most 150 characters.", "danger")
        elif len(phone) > 40 or (
            phone and not re.fullmatch(r"[0-9+().\-\s]+", phone)
        ):
            flash("Enter a valid phone number of at most 40 characters.", "danger")
        elif len(password) < 4:
            flash("Temporary passwords must be at least 4 characters long.", "danger")
        elif len(display_name) > 150:
            flash("Display name must be at most 150 characters.", "danger")
        elif existing_user and (
            existing_user.role == role or existing_user.check_password(password)
        ):
            flash(
                "That username already exists. To update it, both the role and password must change.",
                "danger",
            )
        elif existing_user and existing_user.id == current_user.id:
            flash("Use the profile and password pages to manage your own administrator account.", "danger")
        elif existing_user and existing_user.role == "admin" and role != "admin" and User.query.filter_by(role="admin").count() <= 1:
            flash("The last administrator account cannot be changed to another role.", "danger")
        elif existing_user and (
            Evaluation.query.filter_by(judge_id=existing_user.id).first()
            or Attendance.query.filter_by(marked_by=existing_user.id).first()
            or Attendance.query.join(TeamMember).filter(
                TeamMember.user_id == existing_user.id
            ).first()
            or MentorEvaluation.query.filter(
                db.or_(
                    MentorEvaluation.mentor_id == existing_user.id,
                    MentorEvaluation.reviewed_by == existing_user.id,
                )
            ).first()
            or DomainResultPublication.query.filter_by(
                published_by=existing_user.id
            ).first()
        ):
            flash(
                "This account has evaluation, attendance, or publication history. Keep its role and update its password separately.",
                "danger",
            )
        else:
            team = None
            member_values = None
            assignment_teams = []

            def prepare_account(team_id=None):
                user = existing_user or User(username=username)
                if existing_user:
                    JudgeAssignment.query.filter_by(judge_id=user.id).delete()
                    MentorTeamAssignment.query.filter_by(mentor_id=user.id).delete()
                    CoordinatorTeamAssignment.query.filter_by(
                        coordinator_id=user.id
                    ).delete()
                    CoordinatorDomainAssignment.query.filter_by(
                        coordinator_id=user.id
                    ).delete()
                    TeamMember.query.filter_by(user_id=user.id).update(
                        {"user_id": None}, synchronize_session=False
                    )
                user.role = role
                user.display_name = display_name
                user.email = email
                user.phone = phone
                user.team_id = team_id
                user.set_password(password)
                db.session.add(user)
                return user

            def account_action():
                return "Updated account" if existing_user else "Created user"

            def success_prefix():
                return "Account updated" if existing_user else "Account created"

            if role in {"team_leader", "team_member"}:
                try:
                    team_id = int(request.form.get("team_id", ""))
                except ValueError:
                    team_id = None
                team = db.session.get(Team, team_id) if team_id else None
                if not team:
                    flash("Select the FT team this account belongs to.", "danger")
                elif role == "team_member":
                    full_name = request.form.get("full_name", "").strip()
                    prn = request.form.get("prn", "").strip().upper()
                    existing_prn = TeamMember.query.filter(
                        db.func.upper(TeamMember.prn) == prn
                    ).first()
                    if not full_name or len(full_name) > 150:
                        flash("Enter a member name of at most 150 characters.", "danger")
                    elif not re.fullmatch(r"[A-Z0-9_/-]{3,60}", prn):
                        flash("Student ID must contain 3–60 letters, numbers, underscores, slashes or hyphens.", "danger")
                    elif existing_prn and (
                        not existing_user
                        or existing_prn.user_id != existing_user.id
                    ):
                        flash("That PRN is already registered to a team.", "danger")
                    else:
                        member_values = (full_name, prn)
                if team and (role != "team_member" or member_values):
                    user = prepare_account(team.id)
                    if not display_name:
                        user.display_name = (
                            member_values[0] if member_values else username
                        )
                    if member_values:
                        db.session.add(
                            TeamMember(
                                team_id=team.id,
                                full_name=member_values[0],
                                prn=member_values[1],
                                email=email,
                                phone=phone,
                                user=user,
                            )
                        )
                    log(account_action(), f"{username} ({role}) · {team.team_code}")
                    db.session.commit()
                    flash(f"{success_prefix()} for {team.team_code}.", "success")
            elif role in {"mentor", "coordinator"}:
                try:
                    team_ids = {int(value) for value in request.form.getlist("team_ids")}
                except ValueError:
                    team_ids = set()
                valid_team_count = len(team_ids) in {1, 2}
                if not valid_team_count:
                    requirement = (
                        "Select one or two FT teams for the mentor."
                        if role == "mentor"
                        else "Select one or two FT teams for the coordinator."
                    )
                    flash(requirement, "danger")
                else:
                    assignment_teams = Team.query.filter(Team.id.in_(team_ids)).all()
                    if len(assignment_teams) != len(team_ids):
                        abort(400, description="Invalid FT team assignment.")
                    user = prepare_account()
                    db.session.flush()
                    assignment_model = (
                        MentorTeamAssignment if role == "mentor"
                        else CoordinatorTeamAssignment
                    )
                    assignment_key = "mentor_id" if role == "mentor" else "coordinator_id"
                    db.session.add_all(
                        assignment_model(**{assignment_key: user.id, "team_id": team.id})
                        for team in assignment_teams
                    )
                    log(
                        account_action(),
                        f"{username} ({role}) · "
                        + ", ".join(sorted(team.team_code for team in assignment_teams)),
                    )
                    db.session.commit()
                    flash(
                        f"{success_prefix()} with {len(assignment_teams)} FT team assignment(s).",
                        "success",
                    )
            else:
                user = prepare_account()
                log(account_action(), f"{username} ({role})")
                db.session.commit()
                flash(
                    "Account updated." if existing_user else "User created.",
                    "success",
                )
        return redirect(url_for("main.admin_users"))
    users = User.query.order_by(User.id).all()
    teams = Team.query.order_by(Team.team_code).all()
    team_labels = {team.id: team.team_code for team in teams}
    assigned_team_labels = {}
    for assignment in MentorTeamAssignment.query.all():
        assigned_team_labels.setdefault(assignment.mentor_id, []).append(
            team_labels[assignment.team_id]
        )
    for assignment in CoordinatorTeamAssignment.query.all():
        assigned_team_labels.setdefault(assignment.coordinator_id, []).append(
            team_labels[assignment.team_id]
        )
    return render_template(
        "admin/users.html",
        users=users,
        teams=teams,
        team_labels=team_labels,
        assigned_team_labels=assigned_team_labels,
    )


@main.route("/admin/users/<int:user_id>/contact", methods=["POST"])
@roles_required("admin")
def update_user_contact(user_id):
    user = db.get_or_404(User, user_id)
    email = request.form.get("email", "").strip()
    phone = request.form.get("phone", "").strip()
    if len(email) > 150 or (
        email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)
    ):
        flash("Enter a valid email address of at most 150 characters.", "danger")
        return redirect(url_for("main.admin_users"))
    if len(phone) > 40 or (
        phone and not re.fullmatch(r"[0-9+().\-\s]+", phone)
    ):
        flash("Enter a valid phone number of at most 40 characters.", "danger")
        return redirect(url_for("main.admin_users"))
    user.email = email
    user.phone = phone
    if user.member_record:
        user.member_record.email = email
        user.member_record.phone = phone
    log("Updated user contact details", user.username)
    db.session.commit()
    flash(f"Contact details updated for {user.username}.", "success")
    return redirect(url_for("main.admin_users"))


@main.route("/admin/users/<int:user_id>/toggle", methods=["POST"])
@roles_required("admin")
def toggle_user(user_id):
    user = db.get_or_404(User, user_id)
    if user.id == current_user.id:
        flash("You cannot deactivate your own administrator account.", "danger")
    else:
        user.active = not user.active
        log("Changed user account status", f"{user.username}: {'active' if user.active else 'inactive'}")
        db.session.commit()
        flash(f"Account {user.username} is now {'active' if user.active else 'inactive'}.", "success")
    return redirect(url_for("main.admin_users"))


@main.route("/admin/users/<int:user_id>/delete", methods=["POST"])
@roles_required("admin")
def delete_user(user_id):
    user = db.get_or_404(User, user_id)
    if user.id == current_user.id:
        flash("You cannot delete your own administrator account.", "danger")
        return redirect(url_for("main.admin_users"))
    if Evaluation.query.filter_by(judge_id=user.id).first():
        flash(
            "This account has judge evaluation history. Deactivate it instead to preserve its scores.",
            "danger",
        )
        return redirect(url_for("main.admin_users"))
    if Attendance.query.filter_by(marked_by=user.id).first():
        flash(
            "This account has attendance history. Deactivate it instead to preserve the attendance record.",
            "danger",
        )
        return redirect(url_for("main.admin_users"))
    if (
        MentorEvaluation.query.filter_by(mentor_id=user.id).first()
        or MentorEvaluation.query.filter_by(reviewed_by=user.id).first()
        or DomainResultPublication.query.filter_by(published_by=user.id).first()
    ):
        flash(
            "This account has evaluation or published-results history. Deactivate it instead to preserve the audit trail.",
            "danger",
        )
        return redirect(url_for("main.admin_users"))

    username = user.username
    TeamMember.query.filter_by(user_id=user.id).update(
        {"user_id": None}, synchronize_session=False
    )
    AuditLog.query.filter_by(user_id=user.id).update(
        {"user_id": None}, synchronize_session=False
    )
    UserNotification.query.filter_by(user_id=user.id).delete()
    JudgeAssignment.query.filter_by(judge_id=user.id).delete()
    MentorTeamAssignment.query.filter_by(mentor_id=user.id).delete()
    CoordinatorTeamAssignment.query.filter_by(coordinator_id=user.id).delete()
    CoordinatorDomainAssignment.query.filter_by(coordinator_id=user.id).delete()
    log("Deleted user account", username)
    db.session.delete(user)
    db.session.commit()
    flash(f"Account {username} was deleted.", "success")
    return redirect(url_for("main.admin_users"))


@main.route("/admin/users/<int:user_id>/password", methods=["POST"])
@roles_required("admin")
def reset_user_password(user_id):
    user = db.get_or_404(User, user_id)
    password = request.form.get("password", "")
    if len(password) < 4:
        flash("Passwords must be at least 4 characters long.", "danger")
    else:
        user.set_password(password)
        log("Reset user password", user.username)
        db.session.commit()
        flash(f"Password reset for {user.username}. Share the new password securely.", "success")
    return redirect(url_for("main.admin_users"))


@main.route("/admin/approve/<int:team_id>", methods=["POST"])
@roles_required("admin")
def approve_team(team_id):
    if published_results():
        abort(409, description="Team registration status is frozen after nominations are published.")
    team = db.get_or_404(Team, team_id)
    status = request.form.get("status", "Approved")
    if status not in REGISTRATION_STATUSES:
        abort(400, description="Invalid registration status.")
    team.registration_status = status
    log("Changed registration status", f"{team.team_code}: {team.registration_status}")
    db.session.commit()
    return redirect(url_for("main.admin_teams"))

@main.route("/team")
@roles_required("team_leader")
def team_dashboard():
    team = db.session.get(Team, current_user.team_id)
    if not team:
        abort(403)
    final_results_available = public_leaderboard_published()
    return render_template(
        "team/dashboard.html",
        team=team,
        certificates=certificates_for_user(current_user),
        videos=ProjectVideo.query.filter_by(team_id=team.id)
        .order_by(ProjectVideo.uploaded_at.desc()).all(),
        results=team_results(team) if final_results_available else None,
        domain_rankings=ranked_domain_teams(team.domain_id)
        if team.domain_id and domain_result_published(team.domain_id) else [],
        domain_results_published=bool(
            team.domain_id and domain_result_published(team.domain_id)
        ),
        final_leaderboard_published=final_results_available,
    )

@main.route("/team/update", methods=["POST"])
@roles_required("team_leader")
def team_update():
    team = db.session.get(Team, current_user.team_id)
    if not team: abort(404)
    if published_results():
        abort(409, description="Team details are frozen after nominations are published.")
    team_name = request.form.get("name", team.name or "").strip()
    if not team_name or len(team_name) > 150:
        flash("Team name must contain 1 to 150 characters.", "danger")
        return redirect(url_for("main.team_dashboard"))
    team.name = team_name
    for field in ["leader_name","phone","department","year","project_title","abstract","problem_statement","objectives","technologies"]:
        value = request.form.get(field, getattr(team, field))
        if field in {"abstract", "problem_statement", "objectives", "technologies"}:
            value = str(sanitize_rich_text(value))
        setattr(team, field, value)
    log("Team updated own profile", team.team_code)
    db.session.commit()
    flash("Your team and project details were updated.", "success")
    return redirect(url_for("main.team_dashboard"))


@main.route("/team/members", methods=["POST"])
@roles_required("team_leader")
def register_team_member():
    team = db.session.get(Team, current_user.team_id)
    if not team:
        abort(403)
    if published_results():
        abort(409, description="Team membership is frozen after nominations are published.")

    full_name = request.form.get("full_name", "").strip()
    prn = request.form.get("prn", "").strip().upper()
    roll_number = request.form.get("roll_number", "").strip()
    department = request.form.get("department", "").strip()
    year = request.form.get("year", "").strip()
    email = request.form.get("email", "").strip()
    phone = request.form.get("phone", "").strip()
    if not full_name or len(full_name) > 150:
        flash("Enter a valid member name.", "danger")
        return redirect(url_for("main.team_dashboard"))
    if not re.fullmatch(r"[A-Z0-9_/-]{3,60}", prn):
        flash("Student ID must contain 3–60 letters, numbers, underscores, slashes or hyphens.", "danger")
        return redirect(url_for("main.team_dashboard"))
    if (
        len(roll_number) > 40
        or len(department) > 100
        or len(year) > 30
        or len(email) > 150
        or len(phone) > 40
    ):
        flash("One or more member details are too long. Check the form and try again.", "danger")
        return redirect(url_for("main.team_dashboard"))
    if TeamMember.query.filter(db.func.upper(TeamMember.prn) == prn).first():
        flash("That PRN is already registered to a team.", "danger")
        return redirect(url_for("main.team_dashboard"))

    username_base = "M_" + re.sub(r"[^A-Z0-9]", "", prn)[:65]
    username = username_base
    if User.query.filter_by(username=username).first():
        username = f"{username_base[:65]}_{secrets.token_hex(3).upper()}"
    password = secrets.token_urlsafe(9)
    account = User(
        username=username,
        role="team_member",
        display_name=full_name,
        email=email,
        phone=phone,
        team_id=team.id,
    )
    account.set_password(password)
    member = TeamMember(
        team_id=team.id,
        full_name=full_name,
        prn=prn,
        roll_number=roll_number,
        department=department,
        year=year,
        email=email,
        phone=phone,
        user=account,
    )
    db.session.add(member)
    log("Registered a team member", team.team_code)
    db.session.commit()
    return render_template(
        "team/dashboard.html",
        team=team,
        certificates=certificates_for_user(current_user),
        videos=ProjectVideo.query.filter_by(team_id=team.id)
        .order_by(ProjectVideo.uploaded_at.desc()).all(),
        new_member_credentials={"username": username, "password": password},
    )


@main.route("/team/video", methods=["POST"])
@roles_required("team_leader")
def upload_video():
    team = db.session.get(Team, current_user.team_id)
    if not team:
        abort(403)
    file = request.files.get("video")
    saved_video = save_team_video(team, file)
    if not saved_video:
        return redirect(url_for("main.team_dashboard"))
    log("Uploaded project video", team.team_code)
    db.session.commit()
    remove_replaced_video_files(saved_video[1], saved_video[0])
    flash("Video uploaded.", "success")
    return redirect(url_for("main.team_dashboard"))


@main.route("/team/submissions/<submission_type>", methods=["POST"])
@roles_required("team_leader")
def upload_team_submission(submission_type):
    if submission_type not in TEAM_SUBMISSION_EXTENSIONS:
        abort(404)
    team = db.session.get(Team, current_user.team_id)
    if not team:
        abort(403)
    upload = request.files.get("file")
    if not upload or not upload.filename:
        flash("Choose a file to submit.", "danger")
        return redirect(url_for("main.team_dashboard"))
    saved_submission = save_team_submission(team, submission_type, upload)
    if not saved_submission:
        return redirect(url_for("main.team_dashboard"))
    filename, previous_filename = saved_submission
    log(f"Uploaded project {submission_type}", team.team_code)
    db.session.commit()
    if previous_filename and previous_filename != filename:
        previous_path = os.path.join(
            current_app.config["UPLOAD_FOLDER"], "submissions", previous_filename
        )
        if os.path.basename(previous_filename) == previous_filename and os.path.isfile(previous_path):
            os.remove(previous_path)
    flash(
        f"{'PPT' if submission_type == 'ppt' else 'Project documentation'} submitted.",
        "success",
    )
    return redirect(url_for("main.team_dashboard"))


@main.route("/team/submissions/<int:submission_id>")
@login_required
def download_team_submission(submission_id):
    submission = db.get_or_404(TeamSubmission, submission_id)
    team = db.get_or_404(Team, submission.team_id)
    if not can_view_team(current_user, team):
        abort(403)
    return send_from_directory(
        os.path.join(current_app.config["UPLOAD_FOLDER"], "submissions"),
        submission.filename,
        as_attachment=True,
        download_name=secure_filename(submission.original_name)
        or f"{submission.submission_type}-submission",
    )


@main.route("/team/video/<int:video_id>")
@login_required
def view_project_video(video_id):
    video = db.get_or_404(ProjectVideo, video_id)
    team = db.get_or_404(Team, video.team_id)
    if not can_view_team(current_user, team):
        abort(403)
    return send_from_directory(
        os.path.join(current_app.config["UPLOAD_FOLDER"], "videos"),
        video.filename,
        mimetype="video/mp4",
        as_attachment=False,
    )


@main.route("/member")
@roles_required("team_member")
def member_dashboard():
    team = db.session.get(Team, current_user.team_id)
    if not team:
        abort(403)
    final_results_available = public_leaderboard_published()
    return render_template(
        "generic/dashboard.html",
        heading="Team Member Dashboard",
        team=team,
        member=current_user.member_record,
        certificates=certificates_for_user(current_user),
        videos=ProjectVideo.query.filter_by(team_id=team.id).all(),
        results=team_results(team) if final_results_available else None,
        domain_rankings=ranked_domain_teams(team.domain_id)
        if team.domain_id and domain_result_published(team.domain_id) else [],
        domain_results_published=bool(
            team.domain_id and domain_result_published(team.domain_id)
        ),
        final_leaderboard_published=final_results_available,
        attendance=Attendance.query.filter_by(
            member_id=current_user.member_record.id
        ).order_by(Attendance.event_date.desc()).all() if current_user.member_record else [],
    )


@main.route("/teams/<int:team_id>/workspace")
@login_required
def team_workspace(team_id):
    team = db.session.get(Team, team_id)
    if not team or not can_use_team_workspace(current_user, team):
        abort(404)

    participants = collaboration_users(team)
    eligible_task_participants = task_assignees_for_role(
        current_user.role, participants
    )
    recipient_id = None
    recipient_value = request.args.get("recipient_id")
    if recipient_value:
        try:
            recipient_id = int(recipient_value)
        except ValueError:
            abort(400, description="Invalid chat recipient.")
        if recipient_id == current_user.id or not any(
            user.id == recipient_id for user in participants
        ):
            abort(404)

    if recipient_id is None:
        message_filter = and_(
            TeamChatMessage.team_id == team.id,
            TeamChatMessage.recipient_id.is_(None),
        )
    else:
        message_filter = and_(
            TeamChatMessage.team_id == team.id,
            or_(
                and_(
                    TeamChatMessage.sender_id == current_user.id,
                    TeamChatMessage.recipient_id == recipient_id,
                ),
                and_(
                    TeamChatMessage.sender_id == recipient_id,
                    TeamChatMessage.recipient_id == current_user.id,
                ),
            ),
        )
    messages = (
        TeamChatMessage.query.filter(message_filter)
        .order_by(TeamChatMessage.sent_at.desc(), TeamChatMessage.id.desc())
        .limit(100).all()
    )
    messages.reverse()
    tasks = TeamTask.query.filter_by(team_id=team.id).order_by(
        TeamTask.created_at.desc()
    ).all()
    return render_template(
        "team/workspace.html",
        team=team,
        tasks=tasks,
        project_tasks=WorkTask.query.filter_by(team_id=team.id)
        .order_by(WorkTask.created_at.desc()).all(),
        participants=participants,
        selected_recipient=(
            db.session.get(User, recipient_id) if recipient_id is not None else None
        ),
        messages=messages,
        can_assign_tasks=current_user.role in {
            "team_leader", "mentor", "coordinator"
        },
        member_accounts=task_assignees_for_role(
            current_user.role, eligible_task_participants
        ),
        project_task_users=eligible_task_participants,
        can_create_project_tasks=current_user.role in TASK_CREATOR_ROLES,
    )


def task_page_email_status(task, recipients):
    delivery = send_task_assignment_emails(task, recipients)
    if delivery["failed"]:
        flash(
            f"Task saved and in-app notices sent, but email delivery failed for "
            f"{delivery['failed']} account(s). Check the server mail configuration.",
            "warning",
        )
    elif delivery["not_configured"]:
        flash(
            f"Task saved and in-app notices sent. Email is not configured for "
            f"{delivery['not_configured']} account(s).",
            "warning",
        )
    elif delivery["missing_email"]:
        flash(
            f"Task saved and in-app notices sent. "
            f"{delivery['missing_email']} account(s) have no email address on file.",
            "warning",
        )
    else:
        flash(
            f"Task saved. In-app notices and email were sent to "
            f"{delivery['sent']} account(s).",
            "success",
        )
    sms_delivery = send_task_assignment_sms(task, recipients)
    if sms_delivery["failed"]:
        flash(
            f"Task saved and in-app notices sent, but MSG91 did not accept SMS "
            f"notifications for {sms_delivery['failed']} account(s). Check the "
            "server log and SMS template configuration.",
            "warning",
        )
    elif sms_delivery["not_configured"]:
        flash(
            f"Task saved and in-app notices sent. MSG91 SMS is not configured for "
            f"{sms_delivery['not_configured']} account(s).",
            "warning",
        )
    elif sms_delivery["missing_phone"]:
        flash(
            f"MSG91 accepted SMS notifications for {sms_delivery['accepted']} "
            f"account(s); {sms_delivery['missing_phone']} account(s) have no valid "
            "phone number in their profile.",
            "warning",
        )
    else:
        flash(
            f"MSG91 accepted SMS notifications for {sms_delivery['accepted']} "
            "account(s).",
            "success",
        )


def task_details_message(task):
    deadline = (
        task.due_date.strftime("%d %b %Y")
        if task.due_date else "No deadline set"
    )
    return (
        f"Task: {task.title}\n"
        f"Details: {task.description or 'No additional instructions provided.'}\n"
        f"Deadline: {deadline}"
    )


def post_project_task_group_message(task, action):
    if not task.team_id:
        return
    names = ", ".join(
        recipient.display_name or recipient.username
        for recipient in getattr(task, "_task_recipients", [])
    )
    message_body = (
        f"{action}\n{task_details_message(task)}"
        + (f"\nAssigned to: {names}" if names else "")
    )
    db.session.add(TeamChatMessage(
        team_id=task.team_id,
        sender_id=task.assigner_id,
        body=message_body[:2000],
    ))


def send_task_deadline_reminder(task, task_kind, reminder_date, label, recipients):
    if not recipients:
        return
    try:
        with db.session.begin_nested():
            db.session.add(TaskDeadlineReminder(
                task_kind=task_kind,
                task_id=task.id,
                reminder_date=reminder_date,
            ))
            db.session.flush()
    except IntegrityError:
        return

    due_date = task.due_date.strftime("%d %b %Y")
    title = f"Task deadline {label}: {task.title}"[:200]
    detail = task_details_message(task)
    for recipient in recipients:
        db.session.add(UserNotification(
            user_id=recipient.id,
            title=title,
            body=f"Your task is due {label} ({due_date}).\n{detail}",
        ))
    if task.team_id:
        db.session.add(TeamChatMessage(
            team_id=task.team_id,
            sender_id=task.assigner_id,
            body=(
                f"Automated task deadline reminder: due {label} "
                f"({due_date}).\n{detail}"
            )[:2000],
        ))
    db.session.commit()
    email_delivery = send_task_assignment_emails(
        task, recipients, reminder_label=label
    )
    if email_delivery["failed"]:
        current_app.logger.error(
            "Task deadline reminder email failed for task %s (%s): %s failure(s)",
            task.id,
            label,
            email_delivery["failed"],
        )


def deliver_due_task_reminders():
    today = date.today()
    reminder_dates = (today, today + timedelta(days=1))
    for task in WorkTask.query.filter(WorkTask.due_date.in_(reminder_dates)).all():
        assignments = WorkTaskAssignment.query.filter_by(task_id=task.id).filter(
            WorkTaskAssignment.status != "Completed"
        ).all()
        recipients = [
            assignment.user for assignment in assignments
            if assignment.user and assignment.user.active
        ]
        if recipients:
            label = "tomorrow" if task.due_date == reminder_dates[1] else "today"
            send_task_deadline_reminder(
                task, "work", today, label, recipients
            )
    for task in TeamTask.query.filter(TeamTask.due_date.in_(reminder_dates)).filter(
        TeamTask.status != "Completed"
    ).all():
        if task.assignee and task.assignee.active:
            label = "tomorrow" if task.due_date == reminder_dates[1] else "today"
            send_task_deadline_reminder(
                task, "team", today, label, [task.assignee]
            )


def save_work_task(*, team, recipients, scope):
    title = " ".join(request.form.get("title", "").split())
    description = request.form.get("description", "").strip()
    if not title or len(title) > 200 or len(description) > 4000:
        abort(400, description="Enter a task title and keep details under 4000 characters.")
    if not recipients:
        abort(400, description="Choose at least one task assignee.")
    recipients = list({user.id: user for user in recipients if user.active}.values())
    if not recipients:
        abort(400, description="No active accounts were selected for this task.")
    if any(
        recipient.role not in TASK_ASSIGNABLE_ROLES.get(current_user.role, set())
        for recipient in recipients
    ):
        abort(403, description="You cannot assign a task to one or more selected roles.")
    due_date = None
    due_date_value = request.form.get("due_date", "").strip()
    if due_date_value:
        try:
            due_date = date.fromisoformat(due_date_value)
        except ValueError:
            abort(400, description="Enter a valid task due date.")
    reference_storage_name, reference_original_name = save_task_upload(
        request.files.get("reference_file")
    )
    task = WorkTask(
        scope=scope,
        team_id=team.id if team else None,
        assigner_id=current_user.id,
        title=title,
        description=description,
        due_date=due_date,
        reference_storage_name=reference_storage_name,
        reference_original_name=reference_original_name,
    )
    db.session.add(task)
    db.session.flush()
    task._task_recipients = recipients
    for recipient in recipients:
        db.session.add(WorkTaskAssignment(task_id=task.id, user_id=recipient.id))
        db.session.add(UserNotification(
            user_id=recipient.id,
            title=f"New task: {title}"[:200],
            body=(
                f"{current_user.display_name or current_user.username} assigned you "
                f"a {'project' if team else 'club-wide'} task.\n"
                f"{task_details_message(task)}"
            ),
        ))
    if team:
        post_project_task_group_message(task, "Project task assigned")
    log(
        "Created work task",
        f"{'Project ' + team.team_code if team else 'Club-wide'}: "
        f"{title} assigned to {len(recipients)} account(s)",
    )
    db.session.commit()
    task_page_email_status(task, recipients)
    return task


@main.route("/tasks")
@login_required
def work_tasks():
    assignments = WorkTaskAssignment.query.filter_by(user_id=current_user.id).join(
        WorkTask, WorkTaskAssignment.task_id == WorkTask.id
    ).order_by(WorkTask.created_at.desc()).limit(200).all()
    created_tasks = WorkTask.query.filter_by(assigner_id=current_user.id).order_by(
        WorkTask.created_at.desc()
    ).limit(100).all()
    all_tasks = (
        WorkTask.query.order_by(WorkTask.created_at.desc()).limit(200).all()
        if current_user.role == "admin" else []
    )
    teams = Team.query.order_by(Team.team_code).all() if current_user.role == "admin" else []
    return render_template(
        "tasks/index.html",
        assignments=assignments,
        created_tasks=created_tasks,
        all_tasks=all_tasks,
        teams=teams,
        roles=sorted(TASK_ASSIGNABLE_ROLES["admin"]),
        file_max_mb=TASK_FILE_MAX_BYTES // (1024 * 1024),
        notification_configuration=(
            task_notification_configuration()
            if current_user.role == "admin" else None
        ),
    )


@main.route("/tasks/search-users")
@roles_required("admin")
def search_work_task_users():
    query = request.args.get("q", "").strip()
    if len(query) < 2:
        return jsonify(users=[])
    if len(query) > 80:
        abort(400, description="Search terms cannot exceed 80 characters.")

    scope = request.args.get("scope", "club")
    if scope == "project":
        try:
            team_id = int(request.args.get("team_id", ""))
        except ValueError:
            return jsonify(users=[])
        team = db.session.get(Team, team_id)
        if not team:
            abort(404)
        candidates = task_assignees_for_role(
            "admin", collaboration_users(team)
        )
        matching_users = [
            user for user in candidates
            if query.casefold() in " ".join((
                user.display_name or "",
                user.username or "",
                user.role.replace("_", " "),
            )).casefold()
        ][:21]
    elif scope == "club":
        escaped_query = query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
        search_pattern = f"%{escaped_query}%"
        matching_users = User.query.filter(
            User.active.is_(True),
            User.role.in_(TASK_ASSIGNABLE_ROLES["admin"]),
            or_(
                User.display_name.ilike(search_pattern, escape="\\"),
                User.username.ilike(search_pattern, escape="\\"),
                User.role.ilike(search_pattern, escape="\\"),
            ),
        ).order_by(User.role, User.display_name, User.username).limit(21).all()
    else:
        abort(400, description="Choose a valid task assignment scope.")

    has_more = len(matching_users) > 20
    return jsonify(
        users=[
            {
                "id": user.id,
                "display_name": user.display_name or user.username,
                "username": user.username,
                "role": user.role.replace("_", " ").title(),
            }
            for user in matching_users[:20]
        ],
        has_more=has_more,
    )


def selected_active_users(user_ids):
    selected_ids = set()
    for value in user_ids:
        try:
            selected_ids.add(int(value))
        except (TypeError, ValueError):
            abort(400, description="Invalid task assignee.")
    if not selected_ids:
        return []
    users = User.query.filter(
        User.id.in_(selected_ids),
        User.active.is_(True),
    ).all()
    if {user.id for user in users} != selected_ids:
        abort(400, description="One or more selected accounts are inactive or unavailable.")
    return users


@main.route("/tasks/create", methods=["POST"])
@roles_required("admin")
def create_club_or_admin_project_task():
    scope = request.form.get("scope", "club")
    selected_roles = set(request.form.getlist("target_roles"))
    if selected_roles - USER_ROLES:
        abort(400, description="Choose valid task-assignee roles.")
    recipients = []
    if scope == "club":
        if request.form.get("all_active_users") == "1":
            recipients = User.query.filter(
                User.active.is_(True),
                User.role.in_(TASK_ASSIGNABLE_ROLES["admin"]),
            ).all()
        else:
            if selected_roles:
                if selected_roles - TASK_ASSIGNABLE_ROLES["admin"]:
                    abort(400, description="Admins cannot assign tasks to the selected role.")
                recipients.extend(User.query.filter(
                    User.active.is_(True),
                    User.role.in_(selected_roles),
                ).all())
            selected = selected_active_users(request.form.getlist("assignee_ids"))
            if any(user.role not in TASK_ASSIGNABLE_ROLES["admin"] for user in selected):
                abort(400, description="Admins cannot assign tasks to the selected role.")
            recipients.extend(selected)
        save_work_task(team=None, recipients=recipients, scope="club")
    elif scope == "project":
        try:
            team_id = int(request.form.get("team_id", ""))
        except (TypeError, ValueError):
            abort(400, description="Choose a project team.")
        team = db.session.get(Team, team_id)
        if not team:
            abort(404)
        project_users = task_assignees_for_role(
            "admin", collaboration_users(team)
        )
        allowed_ids = {user.id for user in project_users}
        if selected_roles - TASK_ASSIGNABLE_ROLES["admin"]:
            abort(400, description="Admins cannot assign project tasks to the selected role.")
        if request.form.get("all_project_users") == "1":
            recipients = project_users
        else:
            recipients = [user for user in project_users if user.role in selected_roles]
            selected = selected_active_users(request.form.getlist("assignee_ids"))
            if any(user.id not in allowed_ids for user in selected):
                abort(400, description="Project tasks can only target project participants.")
            recipients.extend(selected)
        save_work_task(team=team, recipients=recipients, scope="project")
    else:
        abort(400, description="Choose a valid task scope.")
    return redirect(url_for("main.work_tasks"))


@main.route("/teams/<int:team_id>/work-tasks", methods=["POST"])
@login_required
def create_project_work_task(team_id):
    team = db.session.get(Team, team_id)
    if (
        current_user.role not in TASK_CREATOR_ROLES
        or not team
        or not can_view_team(current_user, team)
    ):
        abort(404)
    participants = task_assignees_for_role(
        current_user.role, collaboration_users(team)
    )
    if request.form.get("all_project_users") == "1":
        recipients = participants
    else:
        selected = selected_active_users(request.form.getlist("assignee_ids"))
        allowed_ids = {user.id for user in participants}
        if any(user.id not in allowed_ids for user in selected):
            abort(400, description="Project tasks can only target this project’s participants.")
        recipients = selected
    save_work_task(team=team, recipients=recipients, scope="project")
    return redirect(url_for("main.team_workspace", team_id=team.id))


@main.route("/work-tasks/<int:task_id>/status", methods=["POST"])
@login_required
def update_work_task_status(task_id):
    task = db.session.get(WorkTask, task_id)
    assignment = WorkTaskAssignment.query.filter_by(
        task_id=task_id, user_id=current_user.id
    ).first()
    if not task or not assignment:
        abort(404)
    status = request.form.get("status", "")
    if status not in TASK_STATUSES:
        abort(400, description="Choose a valid task status.")
    assignment.status = status
    log("Updated task status", f"{task.title}: {status}")
    db.session.commit()
    flash("Task status updated.", "success")
    return redirect(url_for("main.work_tasks"))


@main.route("/work-tasks/<int:task_id>/submit", methods=["POST"])
@login_required
def submit_work_task_deliverable(task_id):
    task = db.session.get(WorkTask, task_id)
    assignment = WorkTaskAssignment.query.filter_by(
        task_id=task_id, user_id=current_user.id
    ).first()
    if not task or not assignment:
        abort(404)
    storage_name, original_name = save_task_upload(
        request.files.get("deliverable_file")
    )
    if not storage_name:
        abort(400, description="Choose a documentation or code file to submit.")
    assignment.deliverable_storage_name = storage_name
    assignment.deliverable_original_name = original_name
    assignment.submitted_at = datetime.utcnow()
    assignment.status = "Submitted"
    db.session.add(UserNotification(
        user_id=task.assigner_id,
        title=f"Task submitted: {task.title}"[:200],
        body=f"{current_user.display_name or current_user.username} submitted a deliverable.",
    ))
    log("Submitted task deliverable", f"{task.title}: {original_name}")
    db.session.commit()
    flash("Your task deliverable was submitted.", "success")
    return redirect(url_for("main.work_tasks"))


@main.route("/work-tasks/<int:task_id>/files/<kind>", defaults={"assignment_id": None})
@main.route("/work-tasks/<int:task_id>/files/<kind>/<int:assignment_id>")
@login_required
def download_work_task_file(task_id, kind, assignment_id):
    task = db.session.get(WorkTask, task_id)
    if not task:
        abort(404)
    if kind == "reference":
        if (
            current_user.role != "admin"
            and task.assigner_id != current_user.id
            and not WorkTaskAssignment.query.filter_by(
                task_id=task.id, user_id=current_user.id
            ).first()
            and not (task.team and can_view_team(current_user, task.team))
        ):
            abort(404)
        storage_name = task.reference_storage_name
        original_name = task.reference_original_name
    elif kind == "deliverable" and assignment_id is not None:
        task = task_for_user_or_abort(task_id)
        assignment = db.session.get(WorkTaskAssignment, assignment_id)
        if not assignment or assignment.task_id != task.id:
            abort(404)
        if assignment.user_id != current_user.id and (
            current_user.role != "admin" and task.assigner_id != current_user.id
        ):
            abort(404)
        storage_name = assignment.deliverable_storage_name
        original_name = assignment.deliverable_original_name
    else:
        abort(404)
    if not storage_name or not original_name:
        abort(404)
    response = send_file(
        os.path.join(current_app.config["UPLOAD_FOLDER"], "task-files", storage_name),
        as_attachment=True,
        download_name=original_name,
        mimetype="application/octet-stream",
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


@main.route("/teams/<int:team_id>/tasks", methods=["POST"])
@login_required
def create_team_task(team_id):
    team = db.session.get(Team, team_id)
    allowed_assignee_roles = TASK_ASSIGNABLE_ROLES.get(current_user.role, set())
    if (
        not team
        or current_user.role not in TASK_CREATOR_ROLES
        or not allowed_assignee_roles
        or not can_use_team_workspace(current_user, team)
    ):
        abort(404)

    title = request.form.get("title", "").strip()
    description = request.form.get("description", "").strip()
    if not title or len(title) > 200 or len(description) > 4000:
        abort(400, description="Enter a task title and keep the description under 4000 characters.")
    try:
        assignee_id = int(request.form.get("assignee_id", ""))
    except (TypeError, ValueError):
        abort(400, description="Choose an eligible team participant for this task.")
    assignee = db.session.get(User, assignee_id)
    if (
        not assignee
        or assignee.role not in allowed_assignee_roles
        or not any(
            participant.id == assignee.id
            for participant in collaboration_users(team)
        )
        or not assignee.active
    ):
        abort(404)

    due_date = None
    due_date_value = request.form.get("due_date", "").strip()
    if due_date_value:
        try:
            due_date = date.fromisoformat(due_date_value)
        except ValueError:
            abort(400, description="Enter a valid task due date.")
    task = TeamTask(
        team_id=team.id,
        assignee_id=assignee.id,
        assigner_id=current_user.id,
        title=title,
        description=description,
        due_date=due_date,
    )
    db.session.add(task)
    db.session.add(UserNotification(
        user_id=assignee.id,
        title=f"New task: {title}"[:200],
        body=(
            f"{current_user.display_name or current_user.username} assigned you "
            f"a project task for {team.team_code}.\n{task_details_message(task)}"
        ),
    ))
    task._task_recipients = [assignee]
    post_project_task_group_message(task, "Project task assigned")
    log("Assigned team task", f"{team.team_code}: {title} to {assignee.username}")
    db.session.commit()
    task_page_email_status(task, [assignee])
    flash("Task assigned to the team participant.", "success")
    return redirect(url_for("main.team_workspace", team_id=team.id))


@main.route("/tasks/<int:task_id>/status", methods=["POST"])
@login_required
def update_team_task_status(task_id):
    task = db.session.get(TeamTask, task_id)
    if (
        not task
        or task.assignee_id != current_user.id
        or not can_use_team_workspace(current_user, task.team)
    ):
        abort(404)
    status = request.form.get("status", "")
    if status not in {"Open", "In Progress", "Completed"}:
        abort(400, description="Choose a valid task status.")
    task.status = status
    log("Updated team task status", f"{task.team.team_code}: {task.title} · {status}")
    if task.assigner_id != current_user.id:
        db.session.add(UserNotification(
            user_id=task.assigner_id,
            title=f"Task updated: {task.title}"[:200],
            body=(
                f"{current_user.display_name or current_user.username} changed "
                f"the task status to {status}."
            ),
        ))
    db.session.commit()
    flash("Task status updated.", "success")
    return redirect(url_for("main.team_workspace", team_id=task.team_id))


@main.route("/teams/<int:team_id>/chat", methods=["POST"])
@login_required
def send_team_chat_message(team_id):
    team = db.session.get(Team, team_id)
    if not team or not can_use_team_workspace(current_user, team):
        abort(404)
    body = request.form.get("body", "").strip()
    if len(body) > 2000:
        abort(400, description="Chat messages cannot exceed 2000 characters.")
    attachment_storage_name, attachment_original_name = save_chat_attachment()
    if not body and not attachment_storage_name:
        abort(400, description="Write a message or attach a file.")
    recipient_value = request.form.get("recipient_id", "").strip()
    recipient = None
    if recipient_value:
        try:
            recipient_id = int(recipient_value)
        except ValueError:
            abort(400, description="Invalid chat recipient.")
        recipient = next(
            (user for user in collaboration_users(team) if user.id == recipient_id),
            None,
        )
        if not recipient or recipient.id == current_user.id:
            abort(404)
    message = TeamChatMessage(
        team_id=team.id,
        sender_id=current_user.id,
        recipient_id=recipient.id if recipient else None,
        body=body,
        attachment_storage_name=attachment_storage_name,
        attachment_original_name=attachment_original_name,
    )
    db.session.add(message)
    if recipient:
        recipient_ids = {recipient.id}
        history_scope = f"{team.team_code} private chat"
    else:
        recipient_ids = {
            participant.id for participant in collaboration_users(team)
        }
        recipient_ids.update(
            user.id for user in User.query.filter_by(
                role="admin", active=True
            ).all()
        )
        history_scope = f"{team.team_code} group chat"
    notify_chat_message(recipient_ids, current_user, history_scope)
    log("Sent chat message", f"{history_scope} · recipient count {len(recipient_ids - {current_user.id})}")
    db.session.commit()
    if request.accept_mimetypes.best == "application/json":
        return chat_upload_response(message)
    return redirect(url_for(
        "main.team_workspace",
        team_id=team.id,
        recipient_id=recipient.id if recipient else None,
    ))


def team_chat_message_query(team, user, recipient_id):
    if recipient_id is None:
        conversation_filter = TeamChatMessage.recipient_id.is_(None)
    else:
        recipient = next(
            (person for person in collaboration_users(team) if person.id == recipient_id),
            None,
        )
        if not recipient or recipient.id == user.id:
            abort(404)
        conversation_filter = or_(
            and_(
                TeamChatMessage.sender_id == user.id,
                TeamChatMessage.recipient_id == recipient_id,
            ),
            and_(
                TeamChatMessage.sender_id == recipient_id,
                TeamChatMessage.recipient_id == user.id,
            ),
        )
    return TeamChatMessage.query.filter(
        TeamChatMessage.team_id == team.id,
        conversation_filter,
    )


@main.route("/teams/<int:team_id>/chat/messages")
@login_required
def team_chat_messages(team_id):
    team = db.session.get(Team, team_id)
    if not team or not can_use_team_workspace(current_user, team):
        abort(404)
    recipient_value = request.args.get("recipient_id", "").strip()
    try:
        recipient_id = int(recipient_value) if recipient_value else None
    except ValueError:
        abort(400, description="Invalid chat recipient.")
    messages = (
        team_chat_message_query(team, current_user, recipient_id)
        .order_by(TeamChatMessage.sent_at.desc(), TeamChatMessage.id.desc())
        .limit(100).all()
    )
    messages.reverse()
    return {
        "current_user_id": current_user.id,
        "messages": [
            serialize_chat_message(message, message.sender_id == current_user.id)
            for message in messages
        ],
    }


@main.route("/team-chat/<int:message_id>/edit", methods=["POST"])
@login_required
def edit_team_chat_message(message_id):
    message = db.session.get(TeamChatMessage, message_id)
    team = db.session.get(Team, message.team_id) if message else None
    if (
        not message
        or not team
        or message.sender_id != current_user.id
        or not can_use_team_workspace(current_user, team)
    ):
        abort(404)
    if datetime.utcnow() - message.sent_at > CHAT_EDIT_WINDOW:
        abort(409, description="Messages can only be edited within five minutes.")
    if message.recipient_id is not None and message.recipient_id != current_user.id:
        if not any(
            person.id == message.recipient_id
            for person in collaboration_users(team)
        ):
            abort(404)
    body = request.form.get("body", "").strip()
    if len(body) > 2000 or (not body and not message.attachment_storage_name):
        abort(400, description="A message must contain text or an attached file.")
    message.body = body
    message.edited_at = datetime.utcnow()
    db.session.commit()
    return {"message": serialize_chat_message(message, is_own=True)}


@main.route("/team-chat/<int:message_id>/attachment")
@login_required
def download_team_chat_attachment(message_id):
    message = db.session.get(TeamChatMessage, message_id)
    team = db.session.get(Team, message.team_id) if message else None
    if (
        not message
        or not message.attachment_storage_name
        or not team
        or not can_use_team_workspace(current_user, team)
        or (
            message.recipient_id is not None
            and current_user.id not in {message.sender_id, message.recipient_id}
        )
    ):
        abort(404)
    path = os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "chat-files",
        message.attachment_storage_name,
    )
    response = send_file(
        path,
        as_attachment=True,
        download_name=secure_filename(message.attachment_original_name or "") or "attachment",
        mimetype="application/octet-stream",
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


def admin_chat_page(*, administrator, search="", selected_user=None, role_filter=""):
    if administrator:
        users_query = User.query.filter(
            User.active.is_(True),
            User.id != current_user.id,
        )
        if role_filter:
            users_query = users_query.filter_by(role=role_filter)
        if search:
            search_term = f"%{search}%"
            users_query = users_query.filter(or_(
                User.username.ilike(search_term),
                User.display_name.ilike(search_term),
                User.email.ilike(search_term),
                User.role.ilike(search_term),
            ))
        contacts = users_query.order_by(
            User.role, User.display_name, User.username
        ).limit(50).all()
        endpoint = "main.admin_chat_messages"
        send_endpoint = "main.send_admin_chat_message"
        edit_endpoint = "main.edit_admin_chat_message"
        search_endpoint = "main.admin_chats"
        attachment_endpoint = "main.download_admin_chat_attachment"
    else:
        contacts = chat_contacts_for_user(current_user)
        endpoint = "main.user_admin_chat_messages"
        send_endpoint = "main.send_user_admin_chat_message"
        edit_endpoint = "main.edit_user_admin_chat_message"
        search_endpoint = "main.user_chats"
        attachment_endpoint = "main.download_admin_chat_attachment"
    if selected_user and selected_user not in contacts:
        abort(404)
    return render_template(
        "chat/inbox.html",
        administrator=administrator,
        contacts=contacts,
        selected_user=selected_user,
        search=search,
        role_filter=role_filter,
        available_roles=sorted(USER_ROLES),
        search_endpoint=search_endpoint,
        messages_endpoint=endpoint,
        send_endpoint=send_endpoint,
        edit_endpoint=edit_endpoint,
        attachment_endpoint=attachment_endpoint,
    )


@main.route("/admin/chats")
@roles_required("admin")
def admin_chats():
    search = request.args.get("q", "").strip()
    if len(search) > 100:
        abort(400, description="Keep the chat search to 100 characters or fewer.")
    role_filter = request.args.get("role", "").strip()
    if role_filter and role_filter not in USER_ROLES:
        abort(400, description="Invalid chat role filter.")
    user_value = request.args.get("user_id", "").strip()
    try:
        user_id = int(user_value) if user_value else None
    except ValueError:
        abort(400, description="Invalid chat contact.")
    selected_user = active_admin_chat_user(user_id) if user_id else None
    return admin_chat_page(
        administrator=True,
        search=search,
        role_filter=role_filter,
        selected_user=selected_user,
    )


@main.route("/chats")
@login_required
def user_chats():
    if current_user.role == "admin":
        return redirect(url_for("main.admin_chats"))
    user_value = request.args.get("user_id", request.args.get("admin_id", "")).strip()
    try:
        user_id = int(user_value) if user_value else None
    except ValueError:
        abort(400, description="Invalid chat contact.")
    selected_user = active_chat_contact(user_id) if user_id else None
    return admin_chat_page(administrator=False, selected_user=selected_user)


def admin_conversation_query(user_id):
    return AdminChatMessage.query.filter(or_(
        and_(
            AdminChatMessage.sender_id == current_user.id,
            AdminChatMessage.recipient_id == user_id,
        ),
        and_(
            AdminChatMessage.sender_id == user_id,
            AdminChatMessage.recipient_id == current_user.id,
        ),
    ))


def admin_chat_messages_response(user_id):
    messages = admin_conversation_query(user_id).order_by(
        AdminChatMessage.sent_at.desc(), AdminChatMessage.id.desc()
    ).limit(100).all()
    messages.reverse()
    return {
        "current_user_id": current_user.id,
        "messages": [
            serialize_chat_message(message, message.sender_id == current_user.id)
            for message in messages
        ],
    }


@main.route("/admin/chats/<int:user_id>/messages")
@roles_required("admin")
def admin_chat_messages(user_id):
    active_admin_chat_user(user_id)
    return admin_chat_messages_response(user_id)


@main.route("/chats/<int:admin_id>/messages")
@login_required
def user_admin_chat_messages(admin_id):
    if current_user.role == "admin":
        abort(404)
    active_chat_contact(admin_id)
    return admin_chat_messages_response(admin_id)


def send_admin_chat_message(recipient):
    body = request.form.get("body", "").strip()
    if len(body) > 2000:
        abort(400, description="Chat messages cannot exceed 2000 characters.")
    attachment_storage_name, attachment_original_name = save_chat_attachment()
    if not body and not attachment_storage_name:
        abort(400, description="Write a message or attach a file.")
    message = AdminChatMessage(
        sender_id=current_user.id,
        recipient_id=recipient.id,
        body=body,
        attachment_storage_name=attachment_storage_name,
        attachment_original_name=attachment_original_name,
    )
    db.session.add(message)
    notify_chat_message(
        {recipient.id},
        current_user,
        f"private chat with {recipient.username}",
    )
    log("Sent direct message", f"to {recipient.username}")
    db.session.commit()
    return chat_upload_response(message)


@main.route("/admin/chats/<int:user_id>/send", methods=["POST"])
@roles_required("admin")
def send_admin_chat_message_route(user_id):
    recipient = active_admin_chat_user(user_id)
    return send_admin_chat_message(recipient)


@main.route("/chats/<int:admin_id>/send", methods=["POST"])
@login_required
def send_user_admin_chat_message(admin_id):
    if current_user.role == "admin":
        abort(404)
    return send_admin_chat_message(active_chat_contact(admin_id))


def update_admin_direct_message(message_id):
    message = db.session.get(AdminChatMessage, message_id)
    if not message or message.sender_id != current_user.id:
        abort(404)
    if datetime.utcnow() - message.sent_at > CHAT_EDIT_WINDOW:
        abort(409, description="Messages can only be edited within five minutes.")
    body = request.form.get("body", "").strip()
    if len(body) > 2000 or (not body and not message.attachment_storage_name):
        abort(400, description="A message must contain text or an attached file.")
    message.body = body
    message.edited_at = datetime.utcnow()
    db.session.commit()
    return {"message": serialize_chat_message(message, is_own=True)}


@main.route("/admin/chat-messages/<int:message_id>/edit", methods=["POST"])
@roles_required("admin")
def edit_admin_chat_message(message_id):
    return update_admin_direct_message(message_id)


@main.route("/chat-messages/<int:message_id>/edit", methods=["POST"])
@login_required
def edit_user_admin_chat_message(message_id):
    if current_user.role == "admin":
        abort(404)
    return update_admin_direct_message(message_id)


@main.route("/admin/chat-messages/<int:message_id>/attachment")
@login_required
def download_admin_chat_attachment(message_id):
    message = db.session.get(AdminChatMessage, message_id)
    if (
        not message
        or not message.attachment_storage_name
        or current_user.id not in {message.sender_id, message.recipient_id}
    ):
        abort(404)
    path = os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "chat-files",
        message.attachment_storage_name,
    )
    response = send_file(
        path,
        as_attachment=True,
        download_name=secure_filename(message.attachment_original_name or "") or "attachment",
        mimetype="application/octet-stream",
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


def team_results(team):
    judge_evaluations = Evaluation.query.filter_by(
        team_id=team.id, finalized=True, approved=True
    ).all()
    admin_evaluation = AdminEvaluation.query.filter_by(
        team_id=team.id, finalized=True
    ).first()
    if not judge_evaluations or not admin_evaluation:
        return None
    judge_average = sum(item.total for item in judge_evaluations) / len(judge_evaluations)
    return {
        "judge_average": judge_average,
        "admin_score": admin_evaluation.total,
        "total": judge_average + admin_evaluation.total,
        "judge_count": len(judge_evaluations),
        "nomination": Nomination.query.filter_by(
            team_id=team.id, published=True
        ).first(),
        "winner": Winner.query.filter_by(team_id=team.id, published=True).first(),
    }


@main.route("/member/profile", methods=["POST"])
@roles_required("team_member")
def update_member_profile():
    member = current_user.member_record
    if not member:
        abort(403)
    for field, limit in (
        ("full_name", 150),
        ("roll_number", 40),
        ("department", 100),
        ("year", 30),
        ("email", 150),
        ("phone", 40),
    ):
        value = request.form.get(field, getattr(member, field, "") or "").strip()
        if len(value) > limit or (field == "full_name" and not value):
            flash(f"Check the {field.replace('_', ' ')} value and try again.", "danger")
            return redirect(url_for("main.member_dashboard"))
        setattr(member, field, value)
    current_user.display_name = member.full_name
    current_user.email = member.email
    current_user.phone = member.phone
    log("Updated own student profile")
    db.session.commit()
    flash("Your student profile has been updated.", "success")
    return redirect(url_for("main.member_dashboard"))

@main.route("/role/<role>")
@login_required
def simple_dashboard(role):
    allowed = {"Mentor":"mentor", "Coordinator":"coordinator"}
    if allowed.get(role) != current_user.role: abort(403)
    teams = assigned_teams_for_user(current_user)
    domains = sorted(
        {team.domain for team in teams if team.domain},
        key=lambda domain: domain.id,
    )
    results = {
        team.id: team_results(team) for team in teams
        if team.domain_id and domain_result_published(team.domain_id)
    }
    return render_template(
        "generic/dashboard.html",
        heading=f"{role} Dashboard",
        teams=teams,
        team_count=len(teams),
        member_count=sum(len(team.members) for team in teams),
        video_count=sum(
            ProjectVideo.query.filter_by(team_id=team.id).count() for team in teams
        ),
        domains=domains,
        domain_rankings={
            domain.id: ranked_domain_teams(domain.id)
            for domain in domains
            if domain_result_published(domain.id)
        },
        results=results,
        admin_scores={
            team.id: AdminEvaluation.query.filter_by(
                team_id=team.id, finalized=True
            ).first()
            for team in teams
        },
        approved_score_details={
            team.id: approved_team_score_details(team) for team in teams
        },
        mentor_scores={
            team.id: MentorEvaluation.query.filter_by(
                team_id=team.id, mentor_id=current_user.id
            ).first()
            for team in teams
        } if current_user.role == "mentor" else {},
        mentor_proposals={
            team.id: MentorEvaluation.query.filter_by(
                team_id=team.id, status="Approved"
            )
            .order_by(MentorEvaluation.submitted_at.desc()).all()
            for team in teams
        },
        certificates=certificates_for_user(current_user),
    )


@main.route("/attendance", methods=["GET", "POST"])
@roles_required("admin", "coordinator", "judge")
def attendance():
    if current_user.role == "admin":
        teams = Team.query.order_by(Team.team_code).all()
    elif current_user.role == "judge":
        teams = [
            assignment.team
            for assignment in JudgeAssignment.query.filter_by(
                judge_id=current_user.id
            ).join(Team, JudgeAssignment.team_id == Team.id)
            .order_by(Team.team_code)
        ]
    else:
        teams = assigned_teams_for_user(current_user)
    team_ids = {team.id for team in teams}
    if request.method == "POST":
        try:
            event_date = date.fromisoformat(request.form.get("event_date", ""))
        except ValueError:
            abort(400, description="Attendance date must use YYYY-MM-DD format.")
        changed = 0
        for key, status in request.form.items():
            if not key.startswith("status_"):
                continue
            try:
                member_id = int(key.removeprefix("status_"))
            except ValueError:
                abort(400, description="Invalid student identifier.")
            if status not in {"present", "absent", "excused"}:
                abort(400, description="Invalid attendance status.")
            member = db.session.get(TeamMember, member_id)
            if not member or member.team_id not in team_ids:
                abort(403, description="You cannot mark attendance for that student.")
            record = Attendance.query.filter_by(
                member_id=member.id, event_date=event_date
            ).first()
            if record is None:
                record = Attendance(
                    member_id=member.id,
                    event_date=event_date,
                    marked_by=current_user.id,
                )
                db.session.add(record)
            record.status = status
            record.marked_by = current_user.id
            record.marked_at = datetime.utcnow()
            changed += 1
        log("Marked student attendance", f"{event_date.isoformat()}: {changed} students")
        db.session.commit()
        flash(f"Attendance saved for {changed} students.", "success")
        return redirect(url_for("main.attendance", event_date=event_date.isoformat()))

    try:
        selected_date = date.fromisoformat(request.args.get("event_date", date.today().isoformat()))
    except ValueError:
        abort(400, description="Attendance date must use YYYY-MM-DD format.")
    members = (
        TeamMember.query.filter(TeamMember.team_id.in_(team_ids))
        .join(Team)
        .order_by(Team.team_code, TeamMember.full_name)
        .all()
    ) if team_ids else []
    records = Attendance.query.filter(
        Attendance.member_id.in_([member.id for member in members]),
        Attendance.event_date == selected_date,
    ).all() if members else []
    attendance_by_member = {record.member_id: record.status for record in records}
    return render_template(
        "attendance.html",
        teams=teams,
        members=members,
        selected_date=selected_date,
        attendance_by_member=attendance_by_member,
    )

@main.route("/judge")
@roles_required("judge")
def judge_dashboard():
    assignments = (
        JudgeAssignment.query.filter_by(judge_id=current_user.id)
        .join(Team, JudgeAssignment.team_id == Team.id)
        .order_by(Team.team_code)
        .all()
    )
    evaluations = {
        evaluation.team_id: evaluation
        for evaluation in Evaluation.query.filter_by(judge_id=current_user.id).all()
    }
    evaluated_count = sum(
        1 for evaluation in evaluations.values() if evaluation.approved
    )
    pending_count = sum(
        1
        for evaluation in evaluations.values()
        if evaluation.finalized and not evaluation.approved
    )
    submitted_count = sum(
        1 for evaluation in evaluations.values() if evaluation.finalized
    )
    return render_template(
        "judge/dashboard.html",
        assignments=assignments,
        evaluations=evaluations,
        evaluated_count=evaluated_count,
        pending_count=pending_count,
        remaining_count=len(assignments) - submitted_count,
        nominations_published=published_results(),
        certificates=certificates_for_user(current_user),
    )

@main.route("/judge/evaluate/<int:team_id>", methods=["GET","POST"])
@roles_required("judge")
def evaluate(team_id):
    assignment = JudgeAssignment.query.filter_by(
        judge_id=current_user.id, team_id=team_id
    ).first_or_404()
    team = assignment.team
    existing = Evaluation.query.filter_by(team_id=team.id, judge_id=current_user.id).first()
    if request.method == "POST":
        if published_results():
            abort(409, description="Evaluations are frozen after nominations are published.")
        if existing and existing.approved:
            abort(409, description="This evaluation has been approved by the admin and is locked.")
        if existing and existing.finalized:
            abort(409, description="This evaluation is awaiting admin approval.")
        limits = {"problem":8,"innovation":10,"implementation":12,"prototype":15,"impact":5,"presentation":5,"qa":5}
        vals = parse_scores(limits)
        ev = existing or Evaluation(team_id=team.id, judge_id=current_user.id)
        for key, value in vals.items():
            setattr(ev, key, value)
        ev.total = sum(vals.values())
        ev.finalized = True
        ev.approved = False
        ev.approved_by = None
        ev.approved_at = None
        if not existing:
            db.session.add(ev)
        log("Submitted judge evaluation for admin approval", team.team_code)
        db.session.commit()
        flash("Evaluation submitted and is waiting for admin approval.", "success")
        return redirect(url_for("main.judge_dashboard"))
    return render_template(
        "judge/evaluate.html",
        team=team,
        ev=existing,
        nominations_published=published_results(),
    )


@main.route("/judge/evaluate/<int:team_id>/reopen", methods=["POST"])
@roles_required("judge")
def reopen_judge_evaluation(team_id):
    assignment = JudgeAssignment.query.filter_by(
        judge_id=current_user.id, team_id=team_id
    ).first_or_404()
    if published_results():
        abort(409, description="Evaluations cannot be reopened after nominations are published.")
    evaluation = Evaluation.query.filter_by(
        team_id=assignment.team_id, judge_id=current_user.id
    ).first_or_404()
    if evaluation.approved:
        abort(409, description="This evaluation has been approved and cannot be changed.")
    if evaluation.finalized:
        evaluation.finalized = False
        log("Reopened own judge evaluation", assignment.team.team_code)
        db.session.commit()
        flash(
            "Evaluation reopened. Update the marks and submit it again.",
            "warning",
        )
    return redirect(url_for("main.evaluate", team_id=assignment.team_id))


@main.route("/admin/judge-evaluations/<int:evaluation_id>/approve", methods=["POST"])
@roles_required("admin")
def approve_judge_evaluation(evaluation_id):
    evaluation = db.get_or_404(Evaluation, evaluation_id)
    if published_results():
        abort(409, description="Judge scores cannot be approved after nominations are published.")
    if not evaluation.finalized:
        abort(409, description="This judge evaluation has not been submitted.")
    if evaluation.approved:
        evaluation.approved_by = current_user.id
        evaluation.approved_at = datetime.utcnow()
        log(
            "Reconfirmed judge evaluation approval",
            f"{evaluation.team.team_code} · {evaluation.total:.1f}/60 · "
            f"{evaluation.judge.display_name or evaluation.judge.username}",
        )
        db.session.commit()
        flash(
            f"Approval reconfirmed for judge marks for {evaluation.team.team_code}.",
            "success",
        )
        return redirect(url_for("main.admin_overview", section="judge-approved"))
    evaluation.approved = True
    evaluation.approved_by = current_user.id
    evaluation.approved_at = datetime.utcnow()
    log(
        "Approved judge evaluation",
        f"{evaluation.team.team_code} · {evaluation.total:.1f}/60 · "
        f"{evaluation.judge.display_name or evaluation.judge.username}",
    )
    db.session.commit()
    flash(
        f"Judge marks for {evaluation.team.team_code} approved. "
        "The score can now count toward rankings.",
        "success",
    )
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/judge-evaluations/<int:evaluation_id>/edit", methods=["GET", "POST"])
@roles_required("admin")
def edit_judge_evaluation(evaluation_id):
    evaluation = db.get_or_404(Evaluation, evaluation_id)
    if published_results():
        abort(409, description="Judge scores cannot be edited after nominations are published.")
    if request.method == "POST":
        limits = {
            "problem": 8,
            "innovation": 10,
            "implementation": 12,
            "prototype": 15,
            "impact": 5,
            "presentation": 5,
            "qa": 5,
        }
        values = parse_scores(limits)
        for field, score in values.items():
            setattr(evaluation, field, score)
        evaluation.total = sum(values.values())
        evaluation.finalized = True
        evaluation.approved = True
        evaluation.approved_by = current_user.id
        evaluation.approved_at = datetime.utcnow()
        log(
            "Adjusted and approved judge evaluation",
            f"{evaluation.team.team_code}: {evaluation.total:.1f}/60",
        )
        db.session.commit()
        flash(
            f"Adjusted judge marks for {evaluation.team.team_code} and approved them.",
            "success",
        )
        return redirect(url_for("main.admin_overview", section="scores"))
    return render_template("admin/judge_evaluation.html", evaluation=evaluation)


@main.route("/admin/judge-evaluations/<int:evaluation_id>/return", methods=["POST"])
@roles_required("admin")
def return_judge_evaluation(evaluation_id):
    evaluation = db.get_or_404(Evaluation, evaluation_id)
    if published_results():
        abort(409, description="Judge evaluations cannot be changed after nominations are published.")
    if not evaluation.finalized or evaluation.approved:
        abort(409, description="Only a submitted, unapproved evaluation can be returned.")
    evaluation.finalized = False
    evaluation.approved = False
    evaluation.approved_by = None
    evaluation.approved_at = None
    log("Returned judge evaluation for correction", evaluation.team.team_code)
    db.session.commit()
    flash(
        f"Judge marks for {evaluation.team.team_code} were returned for correction.",
        "warning",
    )
    return redirect(url_for("main.admin_dashboard"))


@main.route("/mentor/evaluate/<int:team_id>", methods=["POST"])
@roles_required("mentor")
def submit_mentor_evaluation(team_id):
    assignment = MentorTeamAssignment.query.filter_by(
        mentor_id=current_user.id, team_id=team_id
    ).first_or_404()
    team = assignment.team
    if published_results():
        abort(409, description="Mentor evaluations are frozen after nominations are published.")
    if AdminEvaluation.query.filter_by(
        team_id=team.id, finalized=True
    ).first():
        abort(409, description="The admin has finalized this team's 40-mark evaluation.")
    limits = {
        "development": 10,
        "teamwork": 10,
        "documentation": 8,
        "discipline": 6,
        "feasibility": 6,
    }
    values = parse_scores(limits)
    proposal = MentorEvaluation.query.filter_by(
        team_id=team.id, mentor_id=current_user.id
    ).first()
    if proposal is None:
        proposal = MentorEvaluation(team_id=team.id, mentor_id=current_user.id)
        db.session.add(proposal)
    for field, score in values.items():
        setattr(proposal, field, score)
    proposal.total = sum(values.values())
    proposal.status = "Pending"
    proposal.submitted_at = datetime.utcnow()
    proposal.reviewed_at = None
    proposal.reviewed_by = None
    log("Submitted mentor score proposal", team.team_code)
    db.session.commit()
    flash(f"Your 40-mark proposal for {team.team_code} was sent to the admin for approval.", "success")
    return redirect(url_for("main.simple_dashboard", role="Mentor"))


@main.route("/admin/admin-evaluation/<int:team_id>", methods=["GET","POST"])
@roles_required("admin")
def admin_evaluation(team_id):
    team = db.get_or_404(Team, team_id)
    ev = AdminEvaluation.query.filter_by(team_id=team.id).first()
    if request.method == "POST":
        if published_results():
            abort(409, description="Evaluations are frozen after nominations are published.")
        proposal_id = request.form.get("mentor_evaluation_id", "").strip()
        proposal = None
        if proposal_id:
            try:
                proposal_id = int(proposal_id)
            except ValueError:
                abort(400, description="Invalid mentor evaluation proposal.")
            proposal = MentorEvaluation.query.filter_by(
                id=proposal_id, team_id=team.id, status="Pending"
            ).first_or_404()
        if not ev:
            ev = AdminEvaluation(team_id=team.id); db.session.add(ev)
        limits = {"development":10,"teamwork":10,"documentation":8,"discipline":6,"feasibility":6}
        vals = parse_scores(limits)
        for k,v in vals.items(): setattr(ev,k,v)
        ev.total=sum(vals.values()); ev.finalized=True
        if proposal:
            proposal.status = "Approved"
            proposal.reviewed_by = current_user.id
            proposal.reviewed_at = datetime.utcnow()
            db.session.add(UserNotification(
                user_id=proposal.mentor_id,
                title="Your 40-mark proposal was reviewed",
                body=(
                    f"The admin approved the final score of {ev.total:.1f}/40 "
                    f"for {team.team_code}."
                ),
            ))
        log(
            "Finalized admin evaluation",
            f"{team.team_code}: {ev.total:.1f}/40"
            + (" · mentor proposal reviewed" if proposal else ""),
        )
        db.session.commit()
        flash("Final admin evaluation saved and approved.", "success")
        return redirect(url_for("main.admin_teams"))
    proposals = MentorEvaluation.query.filter_by(
        team_id=team.id
    ).order_by(MentorEvaluation.submitted_at.desc()).all()
    proposal_id = request.args.get("proposal_id", type=int)
    selected_proposal = next(
        (
            item for item in proposals
            if item.id == proposal_id and item.status == "Pending"
        ),
        None,
    )
    return render_template(
        "admin/evaluate.html",
        team=team,
        ev=ev,
        proposals=proposals,
        selected_proposal=selected_proposal,
        form_ev=selected_proposal or ev,
    )


@main.route("/admin/admin-evaluation/<int:team_id>/reopen", methods=["POST"])
@roles_required("admin")
def reopen_admin_evaluation(team_id):
    team = db.get_or_404(Team, team_id)
    evaluation = AdminEvaluation.query.filter_by(team_id=team.id).first_or_404()
    if published_results():
        abort(409, description="Evaluations cannot be reopened after nominations are published.")
    if evaluation.finalized:
        evaluation.finalized = False
        log("Reopened admin evaluation", team.team_code)
        db.session.commit()
        flash("Evaluation reopened. Rankings will exclude it until it is finalized again.", "warning")
    return redirect(url_for("main.admin_evaluation", team_id=team.id))


@main.route("/admin/judge-assignments", methods=["GET", "POST"])
@roles_required("admin")
def judge_assignments():
    judges = User.query.filter_by(role="judge").order_by(User.username).all()
    teams = Team.query.order_by(Team.team_code).all()
    if request.method == "POST":
        if published_results():
            abort(409, description="Judge assignments are frozen after nominations are published.")
        judge = User.query.filter_by(id=request.form.get("judge_id"), role="judge").first_or_404()
        try:
            team_ids = {int(team_id) for team_id in request.form.getlist("team_ids")}
        except ValueError:
            abort(400, description="Invalid team assignment.")
        known_ids = {team.id for team in teams}
        if not team_ids.issubset(known_ids):
            abort(400, description="Invalid team assignment.")

        JudgeAssignment.query.filter_by(judge_id=judge.id).delete()
        db.session.add_all(
            JudgeAssignment(judge_id=judge.id, team_id=team_id)
            for team_id in sorted(team_ids)
        )
        log("Updated judge assignments", f"{judge.username}: {len(team_ids)} teams")
        db.session.commit()
        flash(f"Assignments saved for {judge.display_name or judge.username}.", "success")
        return redirect(url_for("main.judge_assignments"))
    assignment_ids = {
        judge.id: {assignment.team_id for assignment in judge.judge_assignments}
        for judge in judges
    }
    return render_template(
        "admin/judge_assignments.html",
        judges=judges,
        teams=teams,
        assignment_ids=assignment_ids,
    )


@main.route("/admin/role-assignments", methods=["GET"])
@roles_required("admin")
def role_assignments():
    mentors = User.query.filter_by(role="mentor").order_by(User.username).all()
    coordinators = User.query.filter_by(role="coordinator").order_by(User.username).all()
    return render_template(
        "admin/role_assignments.html",
        mentors=mentors,
        coordinators=coordinators,
        teams=Team.query.order_by(Team.team_code).all(),
        mentor_assignments={
            mentor.id: {item.team_id for item in mentor.mentor_team_assignments}
            for mentor in mentors
        },
        coordinator_assignments={
            coordinator.id: {
                item.team_id for item in coordinator.coordinator_team_assignments
            }
            for coordinator in coordinators
        },
    )


@main.route("/admin/role-assignments/mentor", methods=["POST"])
@roles_required("admin")
def save_mentor_assignments():
    mentor = User.query.filter_by(
        id=request.form.get("mentor_id"), role="mentor"
    ).first_or_404()
    try:
        team_ids = {int(value) for value in request.form.getlist("team_ids")}
    except ValueError:
        abort(400, description="Invalid team assignment.")
    if len(team_ids) not in {0, 1, 2}:
        flash("Assign one or two teams, or clear the selection to leave this account unassigned.", "danger")
        return redirect(url_for("main.role_assignments"))
    if not team_ids.issubset({team.id for team in Team.query.all()}):
        abort(400, description="Invalid team assignment.")
    MentorTeamAssignment.query.filter_by(mentor_id=mentor.id).delete()
    db.session.add_all(
        MentorTeamAssignment(mentor_id=mentor.id, team_id=team_id)
        for team_id in sorted(team_ids)
    )
    log("Updated mentor assignments", f"{mentor.username}: {len(team_ids)} teams")
    db.session.commit()
    flash(f"Team assignments saved for {mentor.display_name or mentor.username}.", "success")
    return redirect(url_for("main.role_assignments"))


@main.route("/admin/role-assignments/coordinator", methods=["POST"])
@roles_required("admin")
def save_coordinator_assignments():
    coordinator = User.query.filter_by(
        id=request.form.get("coordinator_id"), role="coordinator"
    ).first_or_404()
    try:
        team_ids = {int(value) for value in request.form.getlist("team_ids")}
    except ValueError:
        abort(400, description="Invalid team assignment.")
    if len(team_ids) not in {0, 1, 2}:
        flash("Assign one or two teams, or clear the selection to leave this account unassigned.", "danger")
        return redirect(url_for("main.role_assignments"))
    if not team_ids.issubset({team.id for team in Team.query.all()}):
        abort(400, description="Invalid team assignment.")
    CoordinatorTeamAssignment.query.filter_by(coordinator_id=coordinator.id).delete()
    db.session.add_all(
        CoordinatorTeamAssignment(coordinator_id=coordinator.id, team_id=team_id)
        for team_id in sorted(team_ids)
    )
    log("Updated coordinator assignments", f"{coordinator.username}: {len(team_ids)} teams")
    db.session.commit()
    flash(f"Team assignments saved for {coordinator.display_name or coordinator.username}.", "success")
    return redirect(url_for("main.role_assignments"))


@main.route("/admin/branding", methods=["GET", "POST"])
@roles_required("admin")
def branding():
    if request.method == "POST":
        college_header = request.form.get("college_header", "").strip()
        if not college_header or len(college_header) > 180:
            flash("Enter a college name of up to 180 characters.", "danger")
            return redirect(url_for("main.branding"))
        uploaded_images = []
        for field, setting_key in (
            ("club_logo", "club_logo"),
            ("college_logo", "college_logo"),
            ("college_banner_image", "college_banner_image"),
        ):
            file = request.files.get(field)
            if file and file.filename:
                if not validate_image_upload(file):
                    flash(
                        f"{field.replace('_', ' ').title()} must be a PNG, JPEG or WebP image under 5 MB.",
                        "danger",
                    )
                    return redirect(url_for("main.branding"))
                uploaded_images.append((file, setting_key))

        save_site_setting("college_header", college_header)
        save_site_setting(
            "college_subtitle", request.form.get("college_subtitle", "").strip()[:100]
        )
        brand_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "branding")
        os.makedirs(brand_directory, exist_ok=True)
        for file, setting_key in uploaded_images:
            extension = os.path.splitext(file.filename)[1].lower()
            filename = secure_filename(f"{setting_key}_{uuid.uuid4().hex}{extension}")
            file.save(os.path.join(brand_directory, filename))
            save_site_setting(setting_key, filename)
        log("Updated exhibition branding")
        db.session.commit()
        flash("College header and club branding saved.", "success")
        return redirect(url_for("main.branding"))
    return render_template("admin/branding.html")


@main.route("/branding/<filename>")
def branding_asset(filename):
    default_assets = {"club_logo.png", "college_logo.jpg"}
    allowed_files = {
        setting.value
        for setting in SiteSetting.query.filter(
            SiteSetting.key.in_(
                ["club_logo", "college_logo", "college_banner_image"]
            )
        ).all()
    } | default_assets
    if filename not in allowed_files:
        abort(404)
    upload_directory = os.path.join(current_app.config["UPLOAD_FOLDER"], "branding")
    if os.path.isfile(os.path.join(upload_directory, filename)):
        return send_from_directory(upload_directory, filename)
    if filename in default_assets:
        return send_from_directory(
            os.path.join(current_app.static_folder, "branding"),
            filename,
        )
    abort(404)


@main.route("/admin/cancel-published-results", methods=["POST"])
@roles_required("admin")
def cancel_published_results():
    if not published_results() and not public_leaderboard_published():
        flash("There are no published nominations or results to cancel.", "warning")
        return redirect(url_for("main.admin_dashboard"))
    domains = Domain.query.all()
    nomination_ids = [item.id for item in Nomination.query.all()]
    for nomination in Nomination.query.filter(Nomination.id.in_(nomination_ids)).all():
        nomination.published = False
    DomainResultPublication.query.delete(synchronize_session=False)
    Winner.query.delete(synchronize_session=False)
    domain_codes = [domain.code for domain in domains]
    announcement_titles = ["Final Winners & Leaderboard Announced"] + [
        f"Domain Results Announced: {code}" for code in domain_codes
    ]
    Announcement.query.filter(
        Announcement.title.in_(announcement_titles)
    ).delete(synchronize_session=False)
    notification_titles = ["Domain nominations published"] + [
        f"{code} results announced" for code in domain_codes
    ]
    UserNotification.query.filter(
        UserNotification.title.in_(notification_titles)
    ).delete(synchronize_session=False)
    save_site_setting("nominations_published", "false")
    save_site_setting("nominations_generated", "false")
    save_site_setting("final_leaderboard_published", "false")
    log(
        "Cancelled published nominations and results",
        "Returned nominations to draft and withdrew domain/final result announcements",
    )
    db.session.commit()
    flash(
        "Published nominations, domain results, and final winner announcement were cancelled. "
        "You can now review marks and generate nominations again.",
        "warning",
    )
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/nominate", methods=["POST"])
@roles_required("admin")
def nominate():
    # Rebuild nominations only before publication; do not silently overwrite a published list.
    if published_results():
        flash("Published nominations are frozen. Contact system administrator for a formal correction.", "warning")
        return redirect(url_for("main.admin_dashboard"))
    ranked_domains = [
        ranked_domain_teams(domain.id)[:3]
        for domain in Domain.query.order_by(Domain.id).all()
    ]
    Nomination.query.delete()
    for ranked in ranked_domains:
        for rank, (team, _) in enumerate(ranked, 1):
            db.session.add(Nomination(team_id=team.id, rank=rank))
    save_site_setting("nominations_generated", "true")
    db.session.commit()
    flash(
        "Top eligible teams ranked and nominations generated. Domains with no eligible teams are left empty; publish when ready.",
        "success",
    )
    return redirect(url_for("main.admin_dashboard"))

@main.route("/admin/publish-nominations", methods=["POST"])
@roles_required("admin")
def publish_nominations():
    generated = db.session.get(SiteSetting, "nominations_generated")
    if not generated or generated.value != "true":
        flash("Generate nominations first.", "danger")
        return redirect(url_for("main.admin_dashboard"))
    nominations = Nomination.query.all()
    nomination_counts = {
        domain.id: sum(
            1 for nomination in nominations
            if nomination.team.domain_id == domain.id
        )
        for domain in Domain.query.all()
    }
    if any(count > 3 for count in nomination_counts.values()):
        abort(409, description="A domain cannot publish more than three nominated teams.")
    for domain_id in nomination_counts:
        if DomainResultPublication.query.filter_by(domain_id=domain_id).first():
            abort(409, description="Nominations cannot be changed after domain results are published.")
    if published_results():
        flash("Nominations are already published and frozen.", "warning")
        return redirect(url_for("main.admin_dashboard"))
    nomination_by_team = {nomination.team_id: nomination for nomination in nominations}
    for nomination in nominations:
        nomination.published = True
    for team in Team.query.filter(Team.domain_id.isnot(None)).all():
        nomination = nomination_by_team.get(team.id)
        if nomination:
            title = "Domain nominations published"
            body = (
                f"{team.team_code} is nominated at domain rank {nomination.rank}. "
                "The admin has published and frozen nominations."
            )
        else:
            title = "Domain nominations published"
            body = (
                f"{team.team_code} was not selected among the top three eligible "
                "teams in its domain. Nominations are now frozen."
            )
        notify_team_stakeholders(
            team,
            title,
            body,
        )
    save_site_setting("nominations_published", "true")
    log(
        "Published nominations",
        f"{len(nominations)} eligible teams across {len(nomination_counts)} domains",
    )
    db.session.commit()
    flash("Nominations published and frozen.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/publish-domain-result/<int:domain_id>", methods=["POST"])
@roles_required("admin")
def publish_domain_result(domain_id):
    domain = db.get_or_404(Domain, domain_id)
    if not published_results():
        abort(409, description="Publish nominations before announcing domain results.")
    if DomainResultPublication.query.filter_by(domain_id=domain.id).first():
        flash(f"Results for {domain.code} have already been announced.", "warning")
        return redirect(url_for("main.admin_dashboard"))

    ranked = ranked_domain_teams(domain.id)
    publication = DomainResultPublication(
        domain_id=domain.id,
        published_by=current_user.id,
    )
    db.session.add(publication)
    announcement_title = f"Domain Results Announced: {domain.code}"
    lines = [f"Official results for {domain.name}:"]
    lines.extend(
        f"Rank {rank}: {team.team_code} — {team.project_title or team.name} "
        f"({score:.2f} / 100)"
        for rank, (team, score) in enumerate(ranked, 1)
    )
    if not ranked:
        lines.append("No eligible fully evaluated teams were available for ranking.")
    db.session.add(Announcement(title=announcement_title, body="\n".join(lines)))
    for team in Team.query.filter_by(domain_id=domain.id).all():
        position = next(
            (
                rank for rank, (ranked_team, _) in enumerate(ranked, 1)
                if ranked_team.id == team.id
            ),
            None,
        )
        detail = (
            f"{team.team_code} placed at domain rank {position}."
            if position else
            f"{team.team_code} was not eligible for the published ranking."
        )
        notify_team_stakeholders(
            team,
            f"{domain.code} results announced",
            f"{detail} View the published domain leaderboard on the Rankings page.",
        )
    log("Announced domain results", domain.code)
    db.session.commit()
    flash(f"Results for {domain.name} are now visible on role dashboards and Rankings.", "success")
    return redirect(url_for("main.admin_dashboard"))

@main.route("/admin/winner", methods=["POST"])
@roles_required("admin")
def select_winner():
    try:
        team_id = int(request.form["team_id"])
    except (KeyError, ValueError):
        abort(400, description="Invalid team.")
    team=db.get_or_404(Team, team_id)
    if not Nomination.query.filter_by(team_id=team.id, published=True).first(): abort(400)
    # one winner per domain
    existing = Winner.query.join(Team, Winner.team_id==Team.id).filter(Team.domain_id==team.domain_id).first()
    if existing:
        flash("A winner is already recorded for this domain.", "warning")
    else:
        db.session.add(Winner(team_id=team.id, published=True))
        log("Selected domain winner", team.team_code)
        db.session.commit()
        flash("Winner selected. Publish the final winners and leaderboard when all five domains are ready.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/publish-leaderboard", methods=["POST"])
@roles_required("admin")
def publish_final_leaderboard():
    domains = Domain.query.order_by(Domain.id).all()
    winners = Winner.query.join(Team, Winner.team_id == Team.id).filter(
        Winner.published.is_(True)
    ).all()
    winner_by_domain = {winner.team.domain_id: winner for winner in winners}
    if len(domains) != 5 or len(winners) != 5 or any(
        domain.id not in winner_by_domain for domain in domains
    ):
        flash(
            "Select exactly one winner in each of the five domains before announcing the final leaderboard.",
            "danger",
        )
        return redirect(url_for("main.admin_dashboard"))
    announcement_title = "Final Winners & Leaderboard Announced"
    announcement = Announcement.query.filter_by(title=announcement_title).first()
    if public_leaderboard_published() and announcement:
        flash("The final winners and leaderboard have already been announced.", "warning")
        return redirect(url_for("main.admin_dashboard"))

    announcement_lines = [
        "The admin has officially announced the FUTURE-TECH INNOVISION 2026 winners:"
    ]
    for domain in domains:
        team = winner_by_domain[domain.id].team
        leader = team.leader_name or team.name or team.team_code
        announcement_lines.append(
            f"{domain.code} — {domain.name}: {team.team_code} · "
            f"{team.project_title or team.name} (Team Leader: {leader})"
        )
    announcement_lines.append("The final leaderboard is now available on the Rankings page.")
    if announcement is None:
        announcement = Announcement(
            title=announcement_title,
            body="\n".join(announcement_lines),
        )
        db.session.add(announcement)
    else:
        announcement.body = "\n".join(announcement_lines)
        announcement.created_at = datetime.utcnow()
    save_site_setting("final_leaderboard_published", "true")
    log("Announced final winners and leaderboard", "One winner for each of five domains")
    db.session.commit()
    flash("Final winners announced and public leaderboard published.", "success")
    return redirect(url_for("main.admin_dashboard"))


@main.route("/admin/certificate-template-starter")
@roles_required("admin")
def certificate_template_starter():
    type_name = request.args.get(
        "type_name", "Certificate of Participation"
    ).strip()
    if type_name not in CERTIFICATE_TYPE_FIELDS:
        abort(400, description="Choose one of the five supported certificate types.")
    filename = secure_filename(type_name.lower().replace(" ", "-")) + "-template.docx"
    return send_file(
        editable_certificate_docx(type_name),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        as_attachment=True,
        download_name=filename,
    )


@main.route("/admin/certificate-spreadsheet-headers", methods=["POST"])
@roles_required("admin")
def certificate_spreadsheet_headers():
    upload = request.files.get("recipient_csv")
    if not upload or not upload.filename:
        return jsonify({"message": "Choose an Excel workbook."}), 400
    if not upload.filename.lower().endswith(".xlsx"):
        return jsonify({"message": "Select an .xlsx Excel workbook."}), 400
    file_data = upload.stream.read(MAX_CERTIFICATE_CSV_SIZE + 1)
    if not file_data or len(file_data) > MAX_CERTIFICATE_CSV_SIZE:
        return jsonify({"message": "The Excel workbook must be no larger than 5 MB."}), 400
    try:
        headers, _ = certificate_excel_rows(file_data)
    except ValueError as error:
        return jsonify({"message": str(error)}), 400
    return jsonify({"headers": headers})


@main.route("/admin/certificate-template/docx-preview", methods=["POST"])
@roles_required("admin")
def certificate_template_docx_preview():
    document_data = read_certificate_template_document_upload(
        request.files.get("template_docx")
    )
    if not document_data:
        abort(400, description="Choose a Word DOCX template to preview.")
    try:
        with zipfile.ZipFile(io.BytesIO(document_data)) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, ET.ParseError):
        abort(400, description="Upload a valid Microsoft Word DOCX document.")
    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    token_pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}|<([a-zA-Z0-9_]+)>")
    aliases = {
        "name": "recipient_name",
        "role": "recipient_role",
        "domain": "domain_name",
        "event": "event_name",
        "day": "event_day",
        "date": "event_date",
        "venue": "event_venue",
        "team": "team_id",
        "title": "project_title",
        "certificate_number": "certificate_id",
        "certificate_no": "certificate_id",
        "qr": "verification_url",
        "verification_url": "verification_url",
    }
    paragraphs = []
    seen_tokens = set()
    for paragraph_index, paragraph in enumerate(
        root.findall(f".//{{{word_ns}}}p")
    ):
        text = "".join(
            node.text or ""
            for node in paragraph.findall(f".//{{{word_ns}}}t")
        ).strip()
        if not text:
            continue
        token_items = []
        for match in token_pattern.finditer(text):
            token = (match.group(1) or match.group(2)).casefold()
            token_items.append({
                "token": token,
                "field": aliases.get(token, token),
            })
            seen_tokens.add(token)
        paragraphs.append({
            "index": paragraph_index,
            "text": text,
            "tokens": token_items,
        })
        if len(paragraphs) >= 150:
            break
    return jsonify({
        "paragraphs": paragraphs,
        "token_count": len(seen_tokens),
    })


@main.route("/admin/certificate-template/docx-preview-pdf", methods=["POST"])
@roles_required("admin")
def certificate_template_docx_preview_pdf():
    document_data = read_certificate_template_document_upload(
        request.files.get("template_docx")
    )
    if not document_data:
        abort(400, description="Choose a Word DOCX template to preview.")
    try:
        pdf_data = convert_certificate_docx_to_pdf(document_data)
    except (RuntimeError, ValueError) as error:
        abort(503, description=str(error))
    return send_file(
        io.BytesIO(pdf_data),
        mimetype="application/pdf",
        download_name="certificate-preview.pdf",
    )


def onlyoffice_document_directory():
    return os.path.join(
        current_app.config["UPLOAD_FOLDER"], "certificate_templates"
    )


def onlyoffice_document_signer():
    return URLSafeTimedSerializer(
        current_app.secret_key,
        salt="onlyoffice-certificate-template-document",
    )


def onlyoffice_url_has_origin(url, origin):
    try:
        candidate = urlsplit(url)
        expected = urlsplit(origin)
        candidate_port = candidate.port or (
            443 if candidate.scheme == "https" else 80
        )
        expected_port = expected.port or (
            443 if expected.scheme == "https" else 80
        )
    except (TypeError, ValueError):
        return False
    return (
        candidate.scheme == expected.scheme
        and candidate.hostname == expected.hostname
        and candidate_port == expected_port
    )


class OnlyOfficeRedirectHandler(urllib.request.HTTPRedirectHandler):
    def __init__(self, expected_origin):
        super().__init__()
        self.expected_origin = expected_origin

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not onlyoffice_url_has_origin(newurl, self.expected_origin):
            raise urllib.error.HTTPError(
                newurl, code, "Cross-origin document redirect rejected.", headers, fp
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


@main.route(
    "/admin/certificate-templates/<int:template_id>/editor-config"
)
@roles_required("admin")
def certificate_template_editor_config(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    server_url = current_app.config.get("ONLYOFFICE_SERVER_URL", "").rstrip("/")
    secret = current_app.config.get("ONLYOFFICE_JWT_SECRET", "")
    if not server_url or not secret:
        return jsonify({
            "message": (
                "The Word editor is not configured. Deploy the ONLYOFFICE service "
                "and set its URL and shared JWT secret."
            )
        }), 503
    if (
        not template.document_file
        or os.path.basename(template.document_file) != template.document_file
    ):
        abort(404, description="This certificate template has no Word document.")
    document_path = os.path.join(
        onlyoffice_document_directory(), template.document_file
    )
    if not os.path.isfile(document_path):
        abort(404, description="The Word source file is missing.")
    editor_key = hashlib.sha256(
        f"{template.id}:{template.document_file}".encode("utf-8")
    ).hexdigest()
    signer = onlyoffice_document_signer()
    document_token = signer.dumps({
        "template_id": template.id,
        "document_file": template.document_file,
    })
    document_path_url = url_for(
        "main.onlyoffice_certificate_template_document",
        template_id=template.id,
        token=document_token,
    )
    public_base_url = (
        current_app.config.get("CERTIFICATE_PUBLIC_BASE_URL")
        or request.url_root.rstrip("/")
    )
    callback_path = url_for(
        "main.onlyoffice_certificate_template_callback",
        template_id=template.id,
        key=editor_key,
    )
    editor_config = {
        "documentType": "word",
        "document": {
            "fileType": "docx",
            "key": editor_key,
            "title": secure_filename(template.certificate_type.name) + ".docx",
            "url": public_base_url + document_path_url,
            "permissions": {
                "edit": True,
                "download": True,
                "print": True,
            },
        },
        "editorConfig": {
            "mode": "edit",
            "callbackUrl": public_base_url + callback_path,
            "user": {
                "id": str(current_user.id),
                "name": current_user.display_name or current_user.username,
            },
            "customization": {
                "autosave": True,
                "forcesave": True,
            },
        },
    }
    token_payload = {
        **editor_config,
        "iat": int(time.time()),
        "exp": int(time.time()) + 3600,
    }
    editor_config["token"] = encode_onlyoffice_token(token_payload, secret)
    return jsonify({
        "server_url": server_url,
        "script_url": server_url + "/web-apps/apps/api/documents/api.js",
        "revision": template.document_file,
        "config": editor_config,
    })


@main.route(
    "/admin/certificate-templates/<int:template_id>/editor-status"
)
@roles_required("admin")
def certificate_template_editor_status(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    return jsonify({"revision": template.document_file})


@main.route(
    "/onlyoffice/certificate-templates/<int:template_id>/document"
)
def onlyoffice_certificate_template_document(template_id):
    token = request.args.get("token", "")
    try:
        document_access = onlyoffice_document_signer().loads(
            token, max_age=8 * 60 * 60
        )
    except BadSignature:
        abort(403, description="The Word editor document link has expired.")
    template = db.get_or_404(CertificateTemplate, template_id)
    document_file = document_access.get("document_file")
    if (
        document_access.get("template_id") != template.id
        or document_file != template.document_file
        or not document_file
        or os.path.basename(document_file) != document_file
    ):
        abort(403, description="The Word editor document link is no longer valid.")
    return send_from_directory(
        onlyoffice_document_directory(),
        document_file,
        as_attachment=False,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )


@main.route(
    "/onlyoffice/certificate-templates/<int:template_id>/callback",
    methods=["POST"],
)
@csrf.exempt
def onlyoffice_certificate_template_callback(template_id):
    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": 1}), 400
    secret = current_app.config.get("ONLYOFFICE_JWT_SECRET", "")
    if not secret:
        current_app.logger.error("ONLYOFFICE callback received without a JWT secret.")
        return jsonify({"error": 1}), 503
    header_token = request.headers.get("Authorization", "")
    if header_token.lower().startswith("bearer "):
        header_token = header_token[7:].strip()
    tokens = [
        candidate
        for candidate in (header_token, body.get("token", ""))
        if isinstance(candidate, str) and candidate
    ]
    valid_callback = False
    for token in dict.fromkeys(tokens):
        try:
            claims = decode_onlyoffice_token(token, secret)
        except (InvalidOnlyOfficeToken, TypeError, ValueError):
            continue
        signed_callback = claims.get("payload", claims)
        if (
            isinstance(signed_callback, dict)
            and str(signed_callback.get("key", "")) == str(body.get("key", ""))
            and str(signed_callback.get("status", ""))
            == str(body.get("status", ""))
        ):
            valid_callback = True
            break
    if not valid_callback or request.args.get("key") != str(body.get("key", "")):
        return jsonify({"error": 1}), 403
    try:
        status = int(body.get("status", 0))
    except (TypeError, ValueError):
        return jsonify({"error": 1}), 400
    if status not in {2, 6}:
        return jsonify({"error": 0})

    template = db.get_or_404(CertificateTemplate, template_id)
    document_url = body.get("url")
    server_url = current_app.config.get("ONLYOFFICE_SERVER_URL", "").rstrip("/")
    if not onlyoffice_url_has_origin(document_url or "", server_url):
        current_app.logger.error(
            "Rejected ONLYOFFICE document callback from an unexpected host."
        )
        return jsonify({"error": 1}), 400
    try:
        opener = urllib.request.build_opener(
            OnlyOfficeRedirectHandler(server_url)
        )
        with opener.open(document_url, timeout=60) as response:
            if not onlyoffice_url_has_origin(response.geturl(), server_url):
                raise ValueError("The document server redirected the saved document.")
            document_data = response.read(MAX_CERTIFICATE_DOCX_SIZE + 1)
    except (urllib.error.URLError, TimeoutError, ValueError) as error:
        current_app.logger.error("Could not retrieve saved ONLYOFFICE document: %s", error)
        return jsonify({"error": 1}), 502
    if not document_data or len(document_data) > MAX_CERTIFICATE_DOCX_SIZE:
        current_app.logger.error("ONLYOFFICE returned an empty or oversized DOCX.")
        return jsonify({"error": 1}), 400
    try:
        with zipfile.ZipFile(io.BytesIO(document_data)) as archive:
            info = archive.getinfo("word/document.xml")
            archive.getinfo("[Content_Types].xml")
            entries = archive.infolist()
            if (
                len(entries) > 2000
                or sum(entry.file_size for entry in entries)
                > MAX_CERTIFICATE_DOCX_SIZE * 4
                or info.file_size > MAX_CERTIFICATE_DOCX_SIZE
            ):
                raise ValueError("The saved DOCX contains oversized content.")
            ET.fromstring(archive.read(info))
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError, ValueError) as error:
        current_app.logger.error("ONLYOFFICE returned an invalid DOCX: %s", error)
        return jsonify({"error": 1}), 400

    positions = json.loads(template.positions_json or "{}")
    converted_background = None
    if positions.get("docx_pdf_primary"):
        try:
            converted_background = convert_certificate_docx_to_pdf(document_data)
        except (RuntimeError, ValueError) as error:
            current_app.logger.error(
                "Could not render saved ONLYOFFICE document: %s", error
            )
            return jsonify({"error": 1}), 422

    directory = onlyoffice_document_directory()
    os.makedirs(directory, exist_ok=True)
    issued_certificates = Certificate.query.filter_by(
        template_id=template.id
    ).all()
    for certificate in issued_certificates:
        if not certificate.template_snapshot_json:
            certificate.template_snapshot_json = json.dumps(
                certificate_template_snapshot(template)
            )
    old_document_file = template.document_file
    old_background_file = template.background_file
    new_document_file = uuid.uuid4().hex + ".docx"
    new_document_path = os.path.join(directory, new_document_file)
    new_background_file = None
    created_paths = [new_document_path]
    try:
        with open(new_document_path, "xb") as saved_document:
            saved_document.write(document_data)
        if converted_background is not None:
            new_background_file = uuid.uuid4().hex + ".pdf"
            new_background_path = os.path.join(directory, new_background_file)
            created_paths.append(new_background_path)
            with open(new_background_path, "xb") as saved_background:
                saved_background.write(converted_background)
            template.page_width, template.page_height = certificate_template_dimensions(
                ".pdf", converted_background
            )
            template.background_file = new_background_file
    except OSError:
        for path in created_paths:
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except OSError:
                current_app.logger.warning(
                    "Could not remove incomplete ONLYOFFICE file %s.",
                    path,
                    exc_info=True,
                )
        current_app.logger.exception("Could not write saved ONLYOFFICE files.")
        return jsonify({"error": 1}), 500
    template.document_file = new_document_file
    template.approved = False
    try:
        db.session.commit()
    except SQLAlchemyError as error:
        db.session.rollback()
        for path in created_paths:
            if os.path.isfile(path):
                os.remove(path)
        current_app.logger.exception("Could not save edited certificate template.")
        return jsonify({"error": 1}), 500
    current_app.logger.info(
        "Certificate template %s was edited in ONLYOFFICE.",
        template.id,
    )
    if not issued_certificates:
        for old_file in (old_document_file, old_background_file if new_background_file else None):
            if old_file and os.path.basename(old_file) == old_file:
                old_path = os.path.join(directory, old_file)
                if os.path.isfile(old_path):
                    try:
                        os.remove(old_path)
                    except OSError:
                        current_app.logger.warning(
                            "Could not remove replaced certificate template file %s.",
                            old_file,
                            exc_info=True,
                        )
    return jsonify({"error": 0})


@main.route("/admin/certificates", methods=["GET","POST"])
@roles_required("admin")
def certificates():
    if request.method == "POST":
        action = request.form.get("action")
        if action in {"upload_files", "upload_zip"}:
            try:
                type_id = int(request.form.get("type_id", ""))
            except ValueError:
                abort(400, description="Select a valid certificate type.")
            certificate_type = db.get_or_404(CertificateType, type_id)
            certificate_uploads = [
                uploaded_file
                for uploaded_file in request.files.getlist("certificate_files")
                if uploaded_file and uploaded_file.filename
            ]
            if not certificate_uploads:
                legacy_archive = request.files.get("certificate_zip")
                if legacy_archive and legacy_archive.filename:
                    certificate_uploads.append(legacy_archive)
            if not certificate_uploads:
                flash("Choose PDF files named Student_ID.pdf or a ZIP of those PDFs.", "danger")
                return redirect(url_for("main.certificates"))
            recipient_scope = request.form.get("pdf_recipient_scope", "matching")
            if recipient_scope == "selected":
                raw_user_ids = request.form.getlist("pdf_recipient_user_ids")
                try:
                    selected_user_ids = {int(value) for value in raw_user_ids}
                except ValueError:
                    abort(400, description="Choose valid student recipients.")
                if not selected_user_ids:
                    flash("Select at least one recipient from the active student list.", "danger")
                    return redirect(url_for("main.certificates"))
                valid_user_ids = {
                    user_id
                    for (user_id,) in db.session.query(User.id)
                    .join(TeamMember, TeamMember.user_id == User.id)
                    .filter(User.active.is_(True), User.id.in_(selected_user_ids))
                    .all()
                }
                if valid_user_ids != selected_user_ids:
                    abort(400, description="One or more selected student accounts are inactive or unavailable.")
            elif recipient_scope == "matching":
                selected_user_ids = None
            else:
                abort(400, description="Choose a valid certificate recipient source.")
            try:
                matches = certificate_files_from_uploads(
                    certificate_uploads,
                    selected_user_ids=selected_user_ids,
                )
            except (
                ValueError,
                zipfile.BadZipFile,
                RuntimeError,
                OSError,
                EOFError,
                NotImplementedError,
            ) as error:
                flash(str(error) or "The certificate ZIP could not be read.", "danger")
                return redirect(url_for("main.certificates"))

            certificate_directory = os.path.join(
                current_app.config["UPLOAD_FOLDER"], "certificates"
            )
            os.makedirs(certificate_directory, exist_ok=True)
            created = 0
            replaced_files = []
            saved_files = []
            try:
                for member, pdf_data in matches:
                    user = member.user
                    certificate = Certificate.query.filter_by(
                        type_id=certificate_type.id,
                        user_id=user.id,
                        dispatch_status="pending",
                    ).order_by(Certificate.issued_at.desc()).first()
                    if certificate is None:
                        certificate = Certificate(
                            certificate_id=new_certificate_id(),
                            type_id=certificate_type.id,
                            recipient_name=member.full_name,
                            recipient_role="Team Member",
                            team_code=member.team.team_code,
                            user_id=user.id,
                        )
                        db.session.add(certificate)
                        db.session.flush()
                        created += 1
                    elif certificate.file_path:
                        replaced_files.append(certificate.file_path)
                    stamped_pdf = add_certificate_verification_to_pdf(
                        pdf_data,
                        certificate.certificate_id,
                    )
                    filename = f"{uuid.uuid4().hex}.pdf"
                    with open(os.path.join(certificate_directory, filename), "xb") as stored_pdf:
                        stored_pdf.write(stamped_pdf)
                    saved_files.append(filename)
                    certificate.recipient_name = member.full_name
                    certificate.recipient_role = "Team Member"
                    certificate.team_code = member.team.team_code
                    certificate.file_path = filename
                    certificate.fields_json = json.dumps({
                        "student_id": member.prn,
                        "username": user.username,
                        "email": member.email,
                    })
                    certificate.dispatch_status = "pending"
                    certificate.dispatched_at = None
                log(
                    "Imported certificates by Student_ID",
                    f"{certificate_type.name}: {len(matches)} matched",
                )
                db.session.commit()
            except Exception:
                db.session.rollback()
                for filename in saved_files:
                    path = os.path.join(certificate_directory, filename)
                    if os.path.isfile(path):
                        os.remove(path)
                raise
            for old_filename in replaced_files:
                if os.path.basename(old_filename) != old_filename:
                    continue
                old_path = os.path.join(certificate_directory, old_filename)
                if os.path.isfile(old_path):
                    os.remove(old_path)
            flash(
                f"Imported {len(matches)} certificate(s) for matching Student IDs. "
                f"{created} new certificate record(s) created.",
                "success",
            )
        elif action == "template":
            event_name = request.form.get("event_name", "").strip()
            event_date = request.form.get("event_date", "").strip()
            event_venue = request.form.get("event_venue", "").strip()
            type_name = request.form.get("type_name", "").strip()
            recipient_role = request.form.get("recipient_role", "")
            if not event_name or len(event_name) > 200:
                abort(400, description="Event name must be 1–200 characters.")
            if not event_date or len(event_date) > 80:
                abort(400, description="Enter an event date or details of at most 80 characters.")
            if not event_venue or len(event_venue) > 200:
                abort(400, description="Enter an event venue of at most 200 characters.")
            if type_name not in CERTIFICATE_TYPE_FIELDS:
                abort(400, description="Select one of the five supported certificate types.")
            if recipient_role not in CERTIFICATE_TEMPLATE_ROLES[type_name]:
                abort(400, description="Select a role supported by this certificate type.")
            document_data = read_certificate_template_document_upload(
                request.files.get("template_docx")
            )
            background_upload = request.files.get("template_file")
            if background_upload and background_upload.filename:
                extension, file_data = read_certificate_template_upload(
                    background_upload
                )
            elif document_data:
                try:
                    file_data = convert_certificate_docx_to_pdf(document_data)
                except (RuntimeError, ValueError) as error:
                    abort(503, description=str(error))
                extension = ".pdf"
            else:
                abort(
                    400,
                    description="Upload a PDF/image background or a Word DOCX template.",
                )
            page_width, page_height = certificate_template_dimensions(
                extension, file_data
            )

            positions = parse_certificate_positions(
                request.form.get("positions_json", "")
            )
            positions["docx_pdf_primary"] = bool(
                document_data
                and not (background_upload and background_upload.filename)
            )
            certificate_type = CertificateType.query.filter_by(name=type_name).first()
            if certificate_type and certificate_type.template:
                abort(409, description="This certificate type already has a template.")
            if certificate_type is None:
                certificate_type = CertificateType(name=type_name)
                db.session.add(certificate_type)
                db.session.flush()
            template_directory = os.path.join(
                current_app.config["UPLOAD_FOLDER"], "certificate_templates"
            )
            os.makedirs(template_directory, exist_ok=True)
            filename = uuid.uuid4().hex + extension
            with open(os.path.join(template_directory, filename), "xb") as stored_template:
                stored_template.write(file_data)
            document_filename = None
            if document_data:
                document_filename = uuid.uuid4().hex + ".docx"
                with open(
                    os.path.join(template_directory, document_filename), "xb"
                ) as stored_document:
                    stored_document.write(document_data)
            template = CertificateTemplate(
                type_id=certificate_type.id,
                event_name=event_name,
                event_date=event_date,
                event_venue=event_venue,
                recipient_role=recipient_role,
                background_file=filename,
                document_file=document_filename,
                positions_json=json.dumps(positions),
                page_width=page_width,
                page_height=page_height,
            )
            db.session.add(template)
            log("Created certificate template", f"{event_name} · {type_name} · {recipient_role}")
            db.session.commit()
            flash(
                "Role-specific certificate template saved as a draft. Review the placeholders and approve it before CSV issuance.",
                "success",
            )
        elif action == "demo_issue":
            try:
                template_id = int(request.form.get("template_id", ""))
                user_id = int(request.form.get("user_id", ""))
            except (TypeError, ValueError):
                abort(400, description="Choose an approved student certificate template and recipient.")
            template = db.get_or_404(CertificateTemplate, template_id)
            if (
                not template.approved
                or template.recipient_role not in {"team_leader", "team_member"}
                or template.certificate_type.name
                not in {
                    "Certificate of Participation",
                    "Certificate of Winner",
                }
            ):
                abort(409, description="Choose an approved participation or winner template for a student.")
            recipient = db.get_or_404(User, user_id)
            if (
                not recipient.active
                or recipient.role != template.recipient_role
            ):
                abort(400, description="Choose an active student account matching the template role.")
            team = db.session.get(Team, recipient.team_id) if recipient.team_id else None
            member = recipient.member_record
            recipient_name = (
                member.full_name if member else recipient.display_name or recipient.username
            )
            team_code = team.team_code if team else ""
            domain_name = team.domain.name if team and team.domain else ""
            project_title = team.project_title if team else ""
            email = user_notification_email(recipient)
            certificate = Certificate(
                certificate_id=new_certificate_id(),
                type_id=template.type_id,
                recipient_name=recipient_name,
                recipient_role=recipient.role.replace("_", " ").title(),
                team_code=team_code[:10],
                user_id=recipient.id,
                template_id=template.id,
                fields_json=json.dumps({
                    "recipient_name": recipient_name,
                    "username": recipient.username,
                    "email": email,
                    "domain_name": domain_name,
                    "team_id": team_code,
                    "project_title": project_title,
                    "event_name": template.event_name,
                    "event_date": template.event_date,
                    "event_venue": template.event_venue,
                    "recipient_role": recipient.role.replace("_", " ").title(),
                    "event_day": "",
                }),
                dispatch_status="dispatched",
                dispatched_at=datetime.utcnow(),
            )
            db.session.add(certificate)
            db.session.add(UserNotification(
                user_id=recipient.id,
                title=f"Demo certificate dispatched: {template.certificate_type.name}"[:200],
                body=(
                    f"A demo certificate for {template.event_name} is available "
                    f"in your dashboard. Certificate number: {certificate.certificate_id}."
                ),
            ))
            log(
                "Generated and dispatched demo certificate",
                f"{certificate.certificate_id} to {recipient.username}",
            )
            db.session.commit()

            if (
                email
                and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email)
                and task_notification_configuration()["email_ready"]
            ):
                delivery = send_certificate_mail_merge_emails(
                    template, [(recipient, email, certificate)]
                )
                if delivery["sent"]:
                    flash(
                        f"Demo certificate sent to {recipient.username} by email and added to Issued Certificates.",
                        "success",
                    )
                else:
                    flash(
                        f"Demo certificate was added to Issued Certificates and the student's dashboard, "
                        f"but email delivery failed. Certificate number: {certificate.certificate_id}.",
                        "warning",
                    )
            else:
                flash(
                    f"Demo certificate dispatched to {recipient.username}'s dashboard and added to "
                    f"Issued Certificates. Email is not configured or no valid email is on the account. "
                    f"Certificate number: {certificate.certificate_id}.",
                    "success",
                )
        elif action == "import_csv":
            try:
                template_id = int(request.form.get("template_id", ""))
            except ValueError:
                abort(400, description="Select a valid approved certificate template.")
            template = db.get_or_404(CertificateTemplate, template_id)
            if not template.approved:
                abort(409, description="Approve this certificate template before issuing certificates.")
            recipient_file = request.files.get("recipient_csv")
            if not recipient_file or not recipient_file.filename:
                abort(400, description="Choose a recipient CSV or Excel file.")
            extension = os.path.splitext(recipient_file.filename)[1].lower()
            file_data = recipient_file.stream.read(MAX_CERTIFICATE_CSV_SIZE + 1)
            if not file_data or len(file_data) > MAX_CERTIFICATE_CSV_SIZE:
                abort(400, description="The recipient file must be no larger than 5 MB.")
            if extension == ".xlsx":
                try:
                    raw_headers, rows = certificate_excel_rows(file_data)
                except ValueError as error:
                    abort(400, description=str(error))
            elif extension == ".csv":
                try:
                    csv_text = file_data.decode("utf-8-sig")
                    reader = csv.DictReader(io.StringIO(csv_text, newline=""))
                    rows = list(reader)
                except (UnicodeDecodeError, csv.Error):
                    abort(400, description="Recipient CSV must be a valid UTF-8 CSV file.")
                raw_headers = reader.fieldnames or []
            else:
                abort(400, description="Recipient file must be a .csv or .xlsx file.")
            if not raw_headers or not rows or len(rows) > MAX_CERTIFICATE_CSV_ROWS:
                abort(400, description="The spreadsheet must have a header and 1–500 recipient rows.")
            if any(
                not isinstance(header, str) or not header.strip()
                for header in raw_headers
            ):
                abort(400, description="Spreadsheet column headers must be non-empty.")
            headers = [
                header.strip().lstrip("\ufeff")
                for header in raw_headers
            ]
            if any(not header for header in headers):
                abort(400, description="Spreadsheet column headers must be non-empty.")
            if len(set(headers)) != len(headers):
                abort(400, description="Spreadsheet column headers must be unique after trimming spaces.")
            if headers != raw_headers:
                rows = [
                    {
                        normalized_header: row.get(raw_header)
                        for raw_header, normalized_header in zip(raw_headers, headers)
                    }
                    for row in rows
                ]

            def mapped_column(field_name, fallback_field_name=None):
                selected = request.form.get(field_name, "").strip()
                if not selected and fallback_field_name:
                    selected = request.form.get(fallback_field_name, "").strip()
                if selected:
                    return selected
                aliases = CERTIFICATE_CSV_COLUMN_ALIASES.get(field_name, ())
                return next(
                    (
                        header
                        for header in headers
                        if re.sub(r"[^a-z0-9]", "", header.casefold()) in aliases
                    ),
                    "",
                )

            username_column = mapped_column("username_column")
            name_column = mapped_column("recipient_name_column")
            team_column = mapped_column("team_id_column", "team_code_column")
            project_column = mapped_column("project_title_column")
            domain_column = mapped_column("domain_column")
            email_column = mapped_column("email_column")
            event_name_column = mapped_column("event_name_column")
            event_date_column = mapped_column("event_date_column")
            event_venue_column = mapped_column("event_venue_column")
            recipient_role_column = mapped_column("recipient_role_column")
            event_day_column = mapped_column("event_day_column")
            required_columns = [username_column, name_column, domain_column]
            template_fields = CERTIFICATE_TYPE_FIELDS.get(
                template.certificate_type.name, ()
            )
            if "team_id" in template_fields:
                required_columns.extend((team_column, project_column))
            optional_columns = (
                team_column,
                project_column,
                email_column,
                event_name_column,
                event_date_column,
                event_venue_column,
                recipient_role_column,
                event_day_column,
            )
            if any(column not in headers for column in required_columns) or any(
                column and column not in headers for column in optional_columns
            ):
                missing_columns = [
                    column
                    for column in required_columns
                    if column not in headers
                ] + [
                    column
                    for column in optional_columns
                    if column and column not in headers
                ]
                abort(
                    400,
                    description=(
                        "The selected CSV column(s) "
                        f"{', '.join(repr(column) for column in missing_columns)} "
                        f"do not match the CSV headers: {', '.join(headers)}. "
                        "Choose a listed header for each required field. Values from each "
                        "selected column are used for that field for every CSV row."
                    ),
                )

            accounts = {
                user.username: user
                for user in User.query.filter_by(
                    role=template.recipient_role, active=True
                ).all()
            }
            seen_usernames = set()
            prepared = []
            for row_number, row in enumerate(rows, 2):
                username = (row.get(username_column) or "").strip()
                recipient_name = (row.get(name_column) or "").strip()
                if not username or not recipient_name:
                    abort(400, description=f"CSV row {row_number} needs a username and a recipient name.")
                if username in seen_usernames:
                    abort(400, description=f"CSV contains duplicate username {username}.")
                seen_usernames.add(username)
                user = accounts.get(username)
                if not user:
                    abort(
                        400,
                        description=(
                            f"CSV row {row_number}: {username} from the mapped account "
                            f"column is not an active {template.recipient_role.replace('_', ' ')} "
                            "username. Map the exact login username (not a serial number or "
                            "database row ID); check User Account Management for account usernames."
                        ),
                    )
                team = db.session.get(Team, user.team_id) if user.team_id else None
                csv_team_code = (row.get(team_column) or "").strip() if team_column else ""
                if csv_team_code:
                    csv_team = Team.query.filter_by(team_code=csv_team_code).first()
                    if not csv_team:
                        abort(400, description=f"CSV row {row_number}: {csv_team_code} is not a valid team ID.")
                    if team and csv_team.id != team.id:
                        abort(400, description=f"CSV row {row_number}: team ID does not match {username}'s assigned team.")
                    if user.role == "mentor" and not MentorTeamAssignment.query.filter_by(
                        mentor_id=user.id, team_id=csv_team.id
                    ).first():
                        abort(400, description=f"CSV row {row_number}: {username} is not assigned to {csv_team_code}.")
                    if user.role == "coordinator" and not CoordinatorTeamAssignment.query.filter_by(
                        coordinator_id=user.id, team_id=csv_team.id
                    ).first():
                        abort(400, description=f"CSV row {row_number}: {username} is not assigned to {csv_team_code}.")
                    team = csv_team
                team_code = team.team_code if team else ""
                project_title = (
                    (row.get(project_column) or "").strip()
                    if project_column else
                    team.project_title if team else ""
                )
                domain_name = (
                    (row.get(domain_column) or "").strip()
                    if domain_column
                    else team.domain.name if team and team.domain else ""
                )
                recipient_email = (
                    (row.get(email_column) or "").strip()
                    if email_column else ""
                ) or user_notification_email(user)
                if not domain_name:
                    abort(400, description=f"CSV row {row_number} needs a domain name.")
                if recipient_email and not re.fullmatch(
                    r"[^@\s]+@[^@\s]+\.[^@\s]+",
                    recipient_email,
                ):
                    abort(400, description=f"Spreadsheet row {row_number} has an invalid email address.")
                if "team_id" in template_fields and not team_code:
                    abort(400, description=f"CSV row {row_number} needs a Team ID.")
                if "project_title" in template_fields and not project_title:
                    abort(400, description=f"CSV row {row_number} needs a project title.")
                event_name = (
                    (row.get(event_name_column) or "").strip()
                    if event_name_column else template.event_name
                )
                event_date = (
                    (row.get(event_date_column) or "").strip()
                    if event_date_column else template.event_date
                )
                event_venue = (
                    (row.get(event_venue_column) or "").strip()
                    if event_venue_column else template.event_venue
                )
                recipient_role = (
                    (row.get(recipient_role_column) or "").strip()
                    if recipient_role_column
                    else user.role.replace("_", " ").title()
                )
                event_day = (
                    (row.get(event_day_column) or "").strip()
                    if event_day_column else ""
                )
                if (
                    len(username) > 80
                    or len(team_code) > 10
                    or len(project_title) > 200
                    or len(domain_name) > 150
                    or len(event_date) > 80
                    or not event_name
                    or len(event_name) > 200
                    or not event_date
                    or not event_venue
                    or len(event_venue) > 200
                    or len(recipient_role) > 80
                    or len(event_day) > 80
                ):
                    abort(400, description=f"CSV row {row_number} contains a field that exceeds its supported length.")
                prepared.append((
                    user,
                    recipient_name,
                    team_code,
                    project_title,
                    domain_name,
                    event_name,
                    event_date,
                    event_venue,
                    recipient_role,
                    event_day,
                    recipient_email,
                ))

            created = 0
            for (
                user,
                recipient_name,
                team_code,
                project_title,
                domain_name,
                event_name,
                event_date,
                event_venue,
                recipient_role,
                event_day,
                recipient_email,
            ) in prepared:
                certificate = Certificate.query.filter_by(
                    type_id=template.type_id,
                    user_id=user.id,
                    dispatch_status="pending",
                ).order_by(Certificate.issued_at.desc()).first()
                if certificate is None:
                    certificate = Certificate(
                        certificate_id=new_certificate_id(),
                        type_id=template.type_id,
                        user_id=user.id,
                        dispatch_status="pending",
                    )
                    db.session.add(certificate)
                    created += 1
                else:
                    certificate.certificate_id = new_certificate_id()
                    certificate.issued_at = datetime.utcnow()
                certificate.recipient_name = recipient_name
                certificate.recipient_role = recipient_role
                certificate.team_code = team_code[:10]
                certificate.template_id = template.id
                certificate.template_snapshot_json = ""
                certificate.dispatched_at = None
                certificate.fields_json = json.dumps({
                    "recipient_name": recipient_name,
                    "username": user.username,
                    "email": recipient_email,
                    "team_id": team_code,
                    "project_title": project_title,
                    "domain_name": domain_name,
                    "event_name": event_name,
                    "event_date": event_date,
                    "event_venue": event_venue,
                    "recipient_role": recipient_role,
                    "event_day": event_day,
                })
            log(
                "Generated pending certificates from mapped spreadsheet",
                f"{template.event_name}: {len(prepared)} {template.recipient_role} recipients; "
                f"source={extension}",
            )
            db.session.commit()
            flash(
                f"Generated {len(prepared)} certificate draft(s); {created} new records. "
                "They remain private until an admin approves and dispatches them.",
                "success",
            )
        elif action == "type":
            name=request.form.get("name", "").strip()
            if not name or len(name) > 150:
                abort(400, description="Certificate type name must be 1–150 characters.")
            if name and not CertificateType.query.filter_by(name=name).first():
                db.session.add(CertificateType(name=name, description=request.form.get("description","")))
                db.session.commit()
                flash("Certificate type created.", "success")
            elif CertificateType.query.filter_by(name=name).first():
                flash("That certificate type already exists.", "warning")
        elif request.form.get("action") == "eligible":
            try:
                type_id = int(request.form["type_id"])
            except (KeyError, ValueError):
                abort(400, description="Invalid certificate type.")
            ctype = db.get_or_404(CertificateType, type_id)
            template = CertificateTemplate.query.filter_by(
                type_id=ctype.id,
                approved=True,
            ).first()
            if not template:
                abort(
                    409,
                    description="Approve a certificate template before issuing certificates.",
                )
            recipients = eligible_certificate_recipients(ctype)
            created = 0
            for user, recipient_name, team_code in recipients:
                existing = Certificate.query.filter_by(
                    type_id=ctype.id, user_id=user.id
                ).first()
                if existing:
                    continue
                team = db.session.get(Team, user.team_id) if user.team_id else None
                db.session.add(Certificate(
                    certificate_id=new_certificate_id(),
                    type_id=ctype.id,
                    recipient_name=recipient_name,
                    recipient_role=user.role.replace("_", " ").title(),
                    team_code=team_code,
                    user_id=user.id,
                    template_id=template.id,
                    fields_json=json.dumps({
                        "domain_name": team.domain.name
                        if team and team.domain
                        else "",
                        "team_id": team_code,
                        "project_title": team.project_title
                        if team
                        else "",
                    }),
                ))
                created += 1
            if created:
                log("Generated eligible certificates", f"{ctype.name}: {created}")
            db.session.commit()
            flash(f"Generated {created} new eligible certificate(s).", "success")
        else:
            try:
                type_id = int(request.form["type_id"])
            except (KeyError, ValueError):
                abort(400, description="Invalid certificate type.")
            ctype=db.get_or_404(CertificateType, type_id)
            recipients=[x.strip() for x in request.form.get("recipients", "").splitlines() if x.strip()]
            if not recipients:
                abort(400, description="Enter at least one recipient name.")
            for name in recipients:
                db.session.add(Certificate(certificate_id=new_certificate_id(),
                    type_id=ctype.id, recipient_name=name, recipient_role=request.form.get("recipient_role",""),
                    team_code=request.form.get("team_code","")))
            log("Generated certificate records", f"{ctype.name}: {len(recipients)}")
            db.session.commit()
            flash(f"Generated {len(recipients)} certificate records.", "success")
        return redirect(url_for("main.certificates"))
    certificates = Certificate.query.order_by(Certificate.issued_at.desc()).all()
    templates = CertificateTemplate.query.order_by(
        CertificateTemplate.created_at.desc()
    ).all()
    approved_templates = [
        template
        for template in templates
        if template.approved
        and template.certificate_type.name in CERTIFICATE_TYPE_FIELDS
    ]
    user_ids = {certificate.user_id for certificate in certificates if certificate.user_id}
    demo_users = User.query.filter(
        User.active.is_(True),
        User.role.in_(("team_leader", "team_member")),
    ).order_by(User.display_name, User.username).all()
    word_certificate_ids = {
        certificate.id
        for certificate in certificates
        if has_certificate_word_source(certificate)
    }
    return render_template(
        "admin/certificates.html",
        types=CertificateType.query.order_by(CertificateType.name).all(),
        templates=templates,
        approved_templates=approved_templates,
        demo_templates=[
            template
            for template in approved_templates
            if template.recipient_role in {"team_leader", "team_member"}
            and template.certificate_type.name
            in {"Certificate of Participation", "Certificate of Winner"}
        ],
        demo_users=demo_users,
        certs=certificates,
        roles=("admin", "team_leader", "team_member", "mentor", "coordinator", "judge"),
        default_positions=CERTIFICATE_TEMPLATE_POSITIONS,
        certificate_types=tuple(CERTIFICATE_TYPE_FIELDS),
        certificate_type_fields=CERTIFICATE_TYPE_FIELDS,
        certificate_field_labels=CERTIFICATE_FIELD_LABELS,
        certificate_role_options={
            name: sorted(roles)
            for name, roles in CERTIFICATE_TEMPLATE_ROLES.items()
        },
        onlyoffice_configured=bool(
            current_app.config.get("ONLYOFFICE_SERVER_URL")
            and current_app.config.get("ONLYOFFICE_JWT_SECRET")
        ),
        user_labels={
            user.id: user.username
            for user in User.query.filter(User.id.in_(user_ids)).all()
        },
        word_certificate_ids=word_certificate_ids,
        dispatch_domains=Domain.query.order_by(Domain.code).all(),
        pdf_recipient_users=(
            User.query.join(TeamMember, TeamMember.user_id == User.id)
            .filter(User.active.is_(True))
            .order_by(User.display_name, User.username)
            .all()
        ),
        dispatch_users=(
            User.query.filter(
                User.id.in_({
                    certificate.user_id
                    for certificate in certificates
                    if certificate.dispatch_status == "pending"
                    and certificate.user_id is not None
                }),
                User.active.is_(True),
            ).order_by(User.display_name, User.username).all()
            if any(
                certificate.dispatch_status == "pending"
                and certificate.user_id is not None
                for certificate in certificates
            )
            else []
        ),
    )


@main.route("/admin/certificates/dispatch", methods=["POST"])
@roles_required("admin")
def dispatch_certificates():
    scope = request.form.get("scope", "")
    pending = Certificate.query.filter_by(dispatch_status="pending")
    if scope == "all":
        certificates = pending.filter(Certificate.user_id.isnot(None)).all()
    elif scope == "domain":
        try:
            domain_id = int(request.form.get("domain_id", ""))
        except (TypeError, ValueError):
            abort(400, description="Choose a valid domain for certificate dispatch.")
        domain = db.get_or_404(Domain, domain_id)
        team_codes = [
            team.team_code for team in Team.query.filter_by(domain_id=domain.id).all()
        ]
        certificates = (
            pending.filter(
                Certificate.user_id.isnot(None),
                Certificate.team_code.in_(team_codes),
            ).all()
            if team_codes else []
        )
    elif scope == "user":
        try:
            user_id = int(request.form.get("user_id", ""))
        except (TypeError, ValueError):
            abort(400, description="Choose an active recipient account.")
        user = db.get_or_404(User, user_id)
        if not user.active:
            abort(409, description="Certificates cannot be dispatched to a deactivated account.")
        certificates = pending.filter(
            Certificate.user_id == user.id
        ).all()
    else:
        abort(400, description="Choose all users, a domain, or one recipient.")

    if not certificates:
        flash("There are no pending certificates for that dispatch selection.", "warning")
        return redirect(url_for("main.certificates"))
    recipients_by_certificate = {}
    invalid_certificates = []
    for certificate in certificates:
        user = db.session.get(User, certificate.user_id)
        if not user or not user.active:
            invalid_certificates.append(certificate.certificate_id)
            continue
        recipients_by_certificate[certificate.id] = user
    if invalid_certificates:
        abort(
            409,
            description=(
                "Some selected recipients have missing or deactivated accounts. "
                "Reactivate or correct these accounts before dispatch: "
                + ", ".join(invalid_certificates[:10])
            ),
        )

    dispatched_at = datetime.utcnow()
    email_items = []
    for certificate in certificates:
        user = recipients_by_certificate[certificate.id]
        certificate.dispatch_status = "dispatched"
        certificate.dispatched_at = dispatched_at
        try:
            fields = json.loads(certificate.fields_json or "{}")
        except (TypeError, json.JSONDecodeError):
            abort(409, description=f"Certificate {certificate.certificate_id} has invalid recipient data.")
        if not isinstance(fields, dict):
            abort(409, description=f"Certificate {certificate.certificate_id} has invalid recipient data.")
        email = fields.get("email") or user_notification_email(user)
        db.session.add(UserNotification(
            user_id=user.id,
            title=f"Certificate dispatched: {certificate.type.name}"[:200],
            body=(
                f"Your certificate has been approved and dispatched. "
                f"Certificate number: {certificate.certificate_id}. "
                "Sign in and open your dashboard to download it."
            ),
        ))
        if email and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
            email_items.append((user, email, certificate))
    log(
        "Approved and dispatched certificates",
        f"scope={scope}; {len(certificates)} certificate(s)",
    )
    db.session.commit()

    email_ready = task_notification_configuration()["email_ready"]
    delivery = {"sent": 0, "failed": 0}
    if email_ready and email_items:
        by_template = {}
        for item in email_items:
            certificate = item[2]
            template = certificate.template
            by_template.setdefault(template.id if template else None, (template, []))[1].append(item)
        for template, recipients in by_template.values():
            result = send_certificate_mail_merge_emails(template, recipients)
            delivery["sent"] += result["sent"]
            delivery["failed"] += result["failed"]
    if delivery["failed"]:
        flash(
            f"Dispatched {len(certificates)} certificate(s) to user dashboards. "
            f"Email sent to {delivery['sent']}; {delivery['failed']} email(s) failed. "
            "The in-app certificates remain available.",
            "warning",
        )
    elif not email_ready:
        flash(
            f"Dispatched {len(certificates)} certificate(s) to user dashboards. "
            "Email is not configured; recipients can download them after signing in.",
            "success",
        )
    else:
        flash(
            f"Approved and dispatched {len(certificates)} certificate(s); "
            f"emailed {delivery['sent']} recipient(s).",
            "success",
        )
    return redirect(url_for("main.certificates"))


@main.route("/admin/certificate-templates/<int:template_id>/edit", methods=["POST"])
@roles_required("admin")
def edit_certificate_template(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    issued_certificates = Certificate.query.filter_by(template_id=template.id).all()
    for certificate in issued_certificates:
        if not certificate.template_snapshot_json:
            certificate.template_snapshot_json = json.dumps(
                certificate_template_snapshot(template)
            )
    event_name = request.form.get("event_name", "").strip()
    event_date = request.form.get("event_date", "").strip()
    event_venue = request.form.get("event_venue", "").strip()
    recipient_role = request.form.get("recipient_role", "")
    if not event_name or len(event_name) > 200:
        abort(400, description="Event name must be 1–200 characters.")
    if not event_date or len(event_date) > 80:
        abort(400, description="Enter an event date or details of at most 80 characters.")
    if not event_venue or len(event_venue) > 200:
        abort(400, description="Enter an event venue of at most 200 characters.")
    type_name = template.certificate_type.name
    if type_name not in CERTIFICATE_TEMPLATE_ROLES or recipient_role not in CERTIFICATE_TEMPLATE_ROLES[type_name]:
        abort(400, description="Select a role supported by this certificate type.")
    template.event_name = event_name
    template.event_date = event_date
    template.event_venue = event_venue
    template.recipient_role = recipient_role
    positions = parse_certificate_positions(
        request.form.get("positions_json", "")
    )
    old_filename = None
    old_document_filename = None
    template_file = request.files.get("template_file")
    if template_file and template_file.filename:
        extension, file_data = read_certificate_template_upload(template_file)
        page_width, page_height = certificate_template_dimensions(
            extension, file_data
        )
        template_directory = os.path.join(
            current_app.config["UPLOAD_FOLDER"], "certificate_templates"
        )
        os.makedirs(template_directory, exist_ok=True)
        old_filename = template.background_file
        filename = uuid.uuid4().hex + extension
        with open(os.path.join(template_directory, filename), "xb") as stored_template:
            stored_template.write(file_data)
        template.background_file = filename
        template.page_width = page_width
        template.page_height = page_height
        positions["docx_pdf_primary"] = False
    template_document = request.files.get("template_docx")
    if template_document and template_document.filename:
        document_data = read_certificate_template_document_upload(template_document)
        converted_background = None
        if positions["docx_pdf_primary"] and not (
            template_file and template_file.filename
        ):
            try:
                converted_background = convert_certificate_docx_to_pdf(
                    document_data
                )
            except (RuntimeError, ValueError) as error:
                abort(503, description=str(error))
        template_directory = os.path.join(
            current_app.config["UPLOAD_FOLDER"], "certificate_templates"
        )
        os.makedirs(template_directory, exist_ok=True)
        old_document_filename = template.document_file
        document_filename = uuid.uuid4().hex + ".docx"
        with open(
            os.path.join(template_directory, document_filename), "xb"
        ) as stored_document:
            stored_document.write(document_data)
        template.document_file = document_filename
        if converted_background is not None:
            page_width, page_height = certificate_template_dimensions(
                ".pdf", converted_background
            )
            old_filename = template.background_file
            filename = uuid.uuid4().hex + ".pdf"
            with open(
                os.path.join(template_directory, filename), "xb"
            ) as stored_template:
                stored_template.write(converted_background)
            template.background_file = filename
            template.page_width = page_width
            template.page_height = page_height
    template.positions_json = json.dumps(positions)
    template.approved = False
    log("Edited certificate template", template.certificate_type.name)
    db.session.commit()
    if (
        old_filename
        and not issued_certificates
        and os.path.basename(old_filename) == old_filename
    ):
        old_path = os.path.join(template_directory, old_filename)
        if os.path.isfile(old_path):
            os.remove(old_path)
    if old_document_filename and os.path.basename(old_document_filename) == old_document_filename:
        old_document_path = os.path.join(template_directory, old_document_filename)
        if not issued_certificates and os.path.isfile(old_document_path):
            os.remove(old_document_path)
    flash("Template changes saved. Approve the updated design before issuing certificates.", "success")
    return redirect(url_for("main.certificates"))


@main.route("/admin/certificate-templates/<int:template_id>/approve", methods=["POST"])
@roles_required("admin")
def approve_certificate_template(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    if os.path.basename(template.background_file) != template.background_file:
        abort(409, description="Certificate template file path is invalid.")
    template_path = os.path.join(
        current_app.config["UPLOAD_FOLDER"],
        "certificate_templates",
        template.background_file,
    )
    if not os.path.isfile(template_path):
        abort(409, description="Upload a valid certificate template before approving it.")
    template.approved = True
    log("Approved certificate template", f"{template.event_name} · {template.certificate_type.name}")
    db.session.commit()
    flash("Certificate template approved for Excel/CSV mail merge and issuance.", "success")
    return redirect(url_for("main.certificates"))


@main.route("/admin/certificate-templates/<int:template_id>/delete", methods=["POST"])
@roles_required("admin")
def delete_certificate_template(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    issued_certificates = Certificate.query.filter_by(template_id=template.id).all()
    template_directory = os.path.join(
        current_app.config["UPLOAD_FOLDER"], "certificate_templates"
    )
    background_filename = template.background_file
    document_filename = template.document_file
    template_type_name = template.certificate_type.name
    for certificate in issued_certificates:
        if not certificate.template_snapshot_json:
            certificate.template_snapshot_json = json.dumps(
                certificate_template_snapshot(template)
            )
        certificate.template_id = None
    template_label = f"{template_type_name} · {template.event_name}"
    db.session.delete(template)
    log(
        "Deleted certificate template",
        f"{template_label}; preserved {len(issued_certificates)} issued certificate(s)",
    )
    db.session.commit()
    if not issued_certificates:
        for filename in (background_filename, document_filename):
            if filename and os.path.basename(filename) == filename:
                path = os.path.join(template_directory, filename)
                if os.path.isfile(path):
                    os.remove(path)
    flash(
        f"Deleted template “{template_type_name}”. "
        + (
            f"{len(issued_certificates)} issued certificate(s) retain their saved design."
            if issued_certificates
            else "Its unused artwork and source file were removed."
        ),
        "success",
    )
    return redirect(url_for("main.certificates"))


@main.route("/admin/certificate-templates/<int:template_id>/background")
@roles_required("admin")
def certificate_template_background_file(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    if os.path.basename(template.background_file) != template.background_file:
        abort(404)
    return send_from_directory(
        os.path.join(current_app.config["UPLOAD_FOLDER"], "certificate_templates"),
        template.background_file,
        as_attachment=False,
    )


@main.route("/admin/certificate-templates/<int:template_id>/document")
@roles_required("admin")
def certificate_template_document_file(template_id):
    template = db.get_or_404(CertificateTemplate, template_id)
    if not template.document_file or os.path.basename(template.document_file) != template.document_file:
        abort(404)
    return send_from_directory(
        os.path.join(current_app.config["UPLOAD_FOLDER"], "certificate_templates"),
        template.document_file,
        as_attachment=True,
        download_name=secure_filename(
            template.certificate_type.name.lower().replace(" ", "-")
        ) + "-template.docx",
    )


@main.route("/admin/certificates/<int:certificate_id>/delete", methods=["POST"])
@roles_required("admin")
def delete_certificate(certificate_id):
    certificate = db.get_or_404(Certificate, certificate_id)
    filename = certificate.file_path
    certificate_label = certificate.certificate_id
    db.session.delete(certificate)
    log("Deleted certificate", certificate_label)
    db.session.commit()
    if filename and os.path.basename(filename) == filename:
        path = os.path.join(
            current_app.config["UPLOAD_FOLDER"], "certificates", filename
        )
        if os.path.isfile(path):
            os.remove(path)
    flash(f"Certificate {certificate_label} was deleted.", "success")
    return redirect(url_for("main.certificates"))


@main.route("/admin/certificates/delete-all", methods=["POST"])
@roles_required("admin")
def delete_all_certificates():
    certificates = Certificate.query.all()
    filenames = [
        certificate.file_path
        for certificate in certificates
        if certificate.file_path
        and os.path.basename(certificate.file_path) == certificate.file_path
    ]
    count = len(certificates)
    for certificate in certificates:
        db.session.delete(certificate)
    log("Deleted all issued certificates", f"{count} certificate record(s)")
    db.session.commit()
    certificate_directory = os.path.join(
        current_app.config["UPLOAD_FOLDER"], "certificates"
    )
    for filename in filenames:
        path = os.path.join(certificate_directory, filename)
        if os.path.isfile(path):
            os.remove(path)
    flash(
        f"Deleted {count} issued certificate record(s). Templates can now be edited.",
        "success",
    )
    return redirect(url_for("main.certificates"))


@main.route("/certificate/<certificate_id>")
def verify_certificate(certificate_id):
    certificate_number = certificate_id.strip().upper()
    cert = Certificate.query.filter_by(certificate_id=certificate_number).first()
    return render_template(
        "public/verify.html",
        cert=cert,
        searched=True,
        certificate_number=certificate_number,
    )


@main.route("/verify-certificate", methods=["GET", "POST"])
def verify_certificate_by_number():
    if request.method == "POST":
        certificate_number = request.form.get("certificate_id", "").strip().upper()
        if not certificate_number or len(certificate_number) > 50:
            flash("Enter a valid certificate number.", "warning")
            return redirect(url_for("main.verify_certificate_by_number"))
        return redirect(
            url_for("main.verify_certificate", certificate_id=certificate_number)
        )
    certificate_number = request.args.get("certificate_id", "").strip().upper()
    if certificate_number:
        cert = Certificate.query.filter_by(
            certificate_id=certificate_number
        ).first()
        return render_template(
            "public/verify.html",
            cert=cert,
            searched=True,
            certificate_number=certificate_number,
        )
    return render_template(
        "public/verify.html",
        cert=None,
        searched=False,
        certificate_number="",
    )


@main.route("/certificate/<certificate_id>/download")
@login_required
def download_certificate(certificate_id):
    cert = Certificate.query.filter_by(certificate_id=certificate_id).first_or_404()
    if not certificate_visible_to_user(cert, current_user):
        abort(403)
    if cert.file_path:
        if os.path.basename(cert.file_path) != cert.file_path:
            abort(404)
        return send_from_directory(
            os.path.join(current_app.config["UPLOAD_FOLDER"], "certificates"),
            cert.file_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=certificate_download_filename(cert, "pdf"),
        )
    if cert.template_id or cert.template_snapshot_json:
        return send_file(
            render_template_certificate(cert),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=certificate_download_filename(cert, "pdf"),
        )
    abort(503, description="This certificate does not have an approved template.")


@main.route("/certificate/<certificate_id>/word-download")
@login_required
def download_certificate_word(certificate_id):
    certificate = Certificate.query.filter_by(
        certificate_id=certificate_id
    ).first_or_404()
    if not certificate_visible_to_user(certificate, current_user):
        abort(403)
    if certificate.template_snapshot_json:
        try:
            snapshot = json.loads(certificate.template_snapshot_json)
        except (TypeError, json.JSONDecodeError):
            abort(503, description="The saved certificate template snapshot is invalid.")
        if not isinstance(snapshot, dict):
            abort(503, description="The saved certificate template snapshot is invalid.")
    elif certificate.template:
        snapshot = certificate_template_snapshot(certificate.template)
    else:
        abort(404, description="This certificate has no saved Word template.")
    document_path, values, saved_positions = certificate_word_merge_data(
        certificate,
        snapshot,
    )
    verification_url = certificate_verification_url(certificate.certificate_id)
    return send_file(
        merged_certificate_docx(
            document_path, values, verification_url, saved_positions
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        as_attachment=True,
        download_name=certificate_download_filename(certificate, "docx"),
    )

@main.route("/admin/announcements", methods=["GET","POST"])
@roles_required("admin")
def announcements():
    if request.method == "POST":
        title = request.form.get("title", "").strip()
        body = request.form.get("body", "").strip()
        if not title or len(title) > 200 or not body:
            flash("Enter an announcement title (up to 200 characters) and details.", "danger")
        else:
            db.session.add(Announcement(title=title, body=body))
            notified_count = 0
            if request.form.get("notify_all_active") == "1":
                active_users = User.query.filter_by(active=True).all()
                db.session.add_all([
                    UserNotification(
                        user_id=user.id,
                        title=title,
                        body=body,
                    )
                    for user in active_users
                ])
                notified_count = len(active_users)
            log("Published announcement", title)
            db.session.commit()
            flash(
                (
                    f"Announcement published and notifications sent to "
                    f"{notified_count} active account(s)."
                    if notified_count
                    else "Announcement published. No in-app notifications were sent."
                ),
                "success",
            )
    return render_template("admin/announcements.html", items=Announcement.query.order_by(Announcement.created_at.desc()).all())


@main.route("/admin/announcements/<int:announcement_id>/delete", methods=["POST"])
@roles_required("admin")
def delete_announcement(announcement_id):
    announcement = db.get_or_404(Announcement, announcement_id)
    title = announcement.title
    db.session.delete(announcement)
    log("Deleted announcement", title)
    db.session.commit()
    flash(f"Announcement “{title}” was deleted.", "success")
    return redirect(url_for("main.announcements"))
