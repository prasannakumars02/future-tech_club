import csv
import io
import json
import os
import re
import sqlite3
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

from openpyxl import Workbook
import pymupdf
import qrcode
from PIL import Image

from app import create_app, ensure_member_schema, seed_initial_data
from app.extensions import db
from app.onlyoffice import decode_onlyoffice_token, encode_onlyoffice_token
from app.routes import certificate_verification_url, merged_certificate_docx
from app.models import (
    AdminEvaluation,
    AdminChatMessage,
    Announcement,
    AuditLog,
    Attendance,
    Certificate,
    CertificateTemplate,
    CertificateType,
    CoordinatorDomainAssignment,
    DomainResultPublication,
    CoordinatorTeamAssignment,
    Domain,
    Evaluation,
    JudgeAssignment,
    MentorTeamAssignment,
    MentorEvaluation,
    Nomination,
    ProjectVideo,
    SiteSetting,
    Team,
    TeamSubmission,
    TeamChatMessage,
    TeamMember,
    TeamTask,
    TaskDeadlineReminder,
    WorkTask,
    WorkTaskAssignment,
    UserNotification,
    User,
    Winner,
)
from scripts.import_student_roster import import_roster, load_roster


class CoreWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.upload_dir = tempfile.TemporaryDirectory()
        cls.app = create_app(
            {
                "TESTING": True,
                "WTF_CSRF_ENABLED": False,
                "SQLALCHEMY_DATABASE_URI": "sqlite://",
                "UPLOAD_FOLDER": cls.upload_dir.name,
                "TEAM_INITIAL_PASSWORD": "ChangeMe123!",
                "BOOTSTRAP_ADMIN_USERNAME": "admin",
                "BOOTSTRAP_ADMIN_PASSWORD": "Admin@12345!",
            }
        )

    @classmethod
    def tearDownClass(cls):
        with cls.app.app_context():
            db.session.remove()
            db.drop_all()
            db.engine.dispose()
        cls.upload_dir.cleanup()

    def setUp(self):
        with self.app.app_context():
            db.drop_all()
            db.create_all()
            seed_initial_data()
        self.client = self.app.test_client()

    def login(self, username, password):
        with self.app.app_context():
            user = User.query.filter_by(username=username).first()
            role = user.role if user else "admin"
        return self.client.post(
            "/login",
            data={"username": username, "password": password, "role": role},
        )

    def test_team_scores_are_hidden_until_final_leaderboard_is_announced(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            team.domain_id = Domain.query.order_by(Domain.id).first().id
            judge = User(username="visibility_judge", role="judge")
            judge.set_password("JudgePassword123!")
            db.session.add(judge)
            db.session.flush()
            db.session.add_all([
                Evaluation(
                    team_id=team.id,
                    judge_id=judge.id,
                    total=52,
                    finalized=True,
                    approved=True,
                ),
                AdminEvaluation(team_id=team.id, total=36, finalized=True),
                DomainResultPublication(
                    domain_id=team.domain_id,
                    published_by=User.query.filter_by(role="admin").first().id,
                ),
            ])
            member = User(
                username="visibility_member",
                role="team_member",
                team_id=team.id,
                display_name="Visible Member",
            )
            member.set_password("MemberPassword123!")
            db.session.add(TeamMember(
                team_id=team.id,
                full_name="Visible Member",
                prn="VISIBILITY001",
                user=member,
            ))
            db.session.commit()

        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        leader_page = self.client.get("/team")
        self.assertEqual(leader_page.status_code, 200)
        self.assertNotIn(b"88.0 / 100", leader_page.data)
        self.assertIn(b"Published ranking", leader_page.data)

        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "visibility_member",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        self.assertNotIn(b"Final score", member_client.get("/member").data)

        with self.app.app_context():
            db.session.add(SiteSetting(
                key="final_leaderboard_published", value="true"
            ))
            db.session.commit()
        leader_page = self.client.get("/team")
        self.assertIn(b"final leaderboard score", leader_page.data)
        self.assertIn(b"88.0 / 100", leader_page.data)
        member_page = member_client.get("/member")
        self.assertIn(b"Final leaderboard results", member_page.data)
        self.assertIn(b"88.0 / 100", member_page.data)

    def test_mentor_sees_only_approved_evaluation_marks(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            mentor = User(
                username="approved_score_mentor",
                role="mentor",
                display_name="Approved Score Mentor",
            )
            mentor.set_password("MentorPassword123!")
            judge = User(
                username="approved_score_judge",
                role="judge",
                display_name="Approved Score Judge",
            )
            judge.set_password("JudgePassword123!")
            pending_mentor = User(username="pending_score_mentor", role="mentor")
            pending_mentor.set_password("PendingMentorPassword123!")
            db.session.add_all([mentor, judge, pending_mentor])
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id),
                Evaluation(
                    team_id=team.id,
                    judge_id=judge.id,
                    total=54,
                    finalized=True,
                    approved=True,
                ),
                AdminEvaluation(team_id=team.id, total=35, finalized=True),
                MentorEvaluation(
                    team_id=team.id,
                    mentor_id=mentor.id,
                    total=32,
                    status="Approved",
                ),
                MentorEvaluation(
                    team_id=team.id,
                    mentor_id=pending_mentor.id,
                    total=17,
                    status="Pending",
                ),
            ])
            db.session.commit()

        self.assertEqual(self.login(
            "approved_score_mentor", "MentorPassword123!"
        ).status_code, 302)
        page = self.client.get("/role/Mentor")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"Approved evaluation marks", page.data)
        self.assertIn(b"Approved Score Judge", page.data)
        self.assertIn(b"54.0 / 60", page.data)
        self.assertIn(b"35.0 / 40", page.data)
        self.assertIn(b"32.0 / 40", page.data)
        self.assertIn(b"89.0 / 100", page.data)
        self.assertNotIn(b"17.0 / 40", page.data)

    def test_team_tasks_and_group_or_direct_chat_are_team_scoped(self):
        with self.app.app_context():
            team_one = Team.query.filter_by(team_code="FT_01").one()
            team_two = Team.query.filter_by(team_code="FT_02").one()
            member_one = User(
                username="workspace_member_one",
                role="team_member",
                team_id=team_one.id,
                display_name="Member One",
            )
            member_one.set_password("MemberPassword123!")
            member_two = User(
                username="workspace_member_two",
                role="team_member",
                team_id=team_two.id,
                display_name="Member Two",
            )
            member_two.set_password("MemberPassword123!")
            leader_one = User(
                username="workspace_leader_one",
                role="team_leader",
                team_id=team_one.id,
                display_name="Leader One",
            )
            leader_one.set_password("LeaderPassword123!")
            leader_two = User(
                username="workspace_leader_two",
                role="team_leader",
                team_id=team_two.id,
                display_name="Leader Two",
            )
            leader_two.set_password("LeaderPassword123!")
            mentor = User(username="workspace_mentor", role="mentor")
            mentor.set_password("MentorPassword123!")
            db.session.add_all([
                member_one, member_two, leader_one, leader_two, mentor
            ])
            db.session.add_all([
                TeamMember(
                    team_id=team_one.id,
                    full_name="Member One",
                    prn="WORKSPACE001",
                    user=member_one,
                ),
                TeamMember(
                    team_id=team_two.id,
                    full_name="Member Two",
                    prn="WORKSPACE002",
                    user=member_two,
                ),
            ])
            db.session.flush()
            db.session.add(MentorTeamAssignment(
                mentor_id=mentor.id, team_id=team_one.id
            ))
            db.session.commit()
            team_one_id = team_one.id
            member_one_id = member_one.id
            member_two_id = member_two.id
            leader_one_id = leader_one.id
            leader_two_id = leader_two.id
            mentor_id = mentor.id

        mentor_client = self.app.test_client()
        self.assertEqual(mentor_client.post("/login", data={
            "username": "workspace_mentor",
            "password": "MentorPassword123!",
            "role": "mentor",
        }).status_code, 302)
        response = mentor_client.post(
            f"/teams/{team_one_id}/tasks",
            data={
                "title": "Prepare demo",
                "description": "Check the live prototype.",
                "assignee_id": member_one_id,
            },
        )
        self.assertEqual(response.status_code, 302)
        workspace_page = mentor_client.get(f"/teams/{team_one_id}/workspace")
        self.assertIn(b"Leader One", workspace_page.data)
        self.assertIn(b"Member One", workspace_page.data)
        self.assertIn(b"Assign to team leader or member", workspace_page.data)
        leader_task_response = mentor_client.post(
            f"/teams/{team_one_id}/tasks",
            data={
                "title": "Review team plan",
                "assignee_id": leader_one_id,
            },
        )
        self.assertEqual(leader_task_response.status_code, 302)
        with self.app.app_context():
            self.assertTrue(TeamTask.query.filter_by(
                team_id=team_one_id,
                assignee_id=leader_one_id,
                title="Review team plan",
            ).first())
            self.assertTrue(UserNotification.query.filter_by(
                user_id=leader_one_id,
                title="New task: Review team plan",
            ).first())
        self.assertEqual(mentor_client.post(
            f"/teams/{team_one_id}/tasks",
            data={
                "title": "Reject other team leader",
                "assignee_id": leader_two_id,
            },
        ).status_code, 404)
        mentor_client.post(
            f"/teams/{team_one_id}/chat", data={"body": "Hello team"}
        )
        mentor_client.post(
            f"/teams/{team_one_id}/chat",
            data={"body": "Private note", "recipient_id": str(member_one_id)},
        )
        with self.app.app_context():
            task_id = TeamTask.query.filter_by(title="Prepare demo").one().id
            self.assertEqual(TeamChatMessage.query.count(), 4)
            member_notifications = UserNotification.query.filter_by(
                user_id=member_one_id
            ).all()
            self.assertTrue(any(
                item.title == "New chat message"
                and "Private note" not in item.body
                and "mentor" in item.body.lower()
                for item in member_notifications
            ))
            self.assertTrue(any(
                item.title == "New chat message" and "group chat" in item.body
                for item in member_notifications
            ))
            self.assertTrue(any(
                item.title == "New task: Prepare demo"
                for item in member_notifications
            ))

        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "workspace_member_one",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        group_page = member_client.get(f"/teams/{team_one_id}/chat/messages")
        self.assertTrue(any(
            message["body"] == "Hello team"
            for message in group_page.json["messages"]
        ))
        direct_page = member_client.get(
            f"/teams/{team_one_id}/chat/messages?recipient_id={mentor_id}"
        )
        self.assertEqual(direct_page.json["messages"][0]["body"], "Private note")
        self.assertEqual(member_client.post(
            f"/tasks/{task_id}/status", data={"status": "Completed"}
        ).status_code, 302)
        with self.app.app_context():
            self.assertEqual(db.session.get(TeamTask, task_id).status, "Completed")

        other_client = self.app.test_client()
        self.assertEqual(other_client.post("/login", data={
            "username": "workspace_member_two",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        self.assertEqual(
            other_client.get(f"/teams/{team_one_id}/workspace").status_code, 404
        )
        self.assertEqual(other_client.post(
            f"/tasks/{task_id}/status", data={"status": "Completed"}
        ).status_code, 404)
        self.assertEqual(other_client.post(
            f"/teams/{team_one_id}/chat", data={"body": "Not allowed"}
        ).status_code, 404)
        with self.app.app_context():
            self.assertEqual(TeamChatMessage.query.count(), 4)
            self.assertFalse(any(
                item.title == "New chat message"
                and "Not allowed" in item.body
                for item in UserNotification.query.all()
            ))
            self.assertNotEqual(
                db.session.get(User, member_two_id).team_id, team_one_id
            )

    def test_project_and_club_tasks_support_role_assignments_and_file_submissions(self):
        with self.app.app_context():
            team_one = Team.query.filter_by(team_code="FT_01").one()
            team_two = Team.query.filter_by(team_code="FT_02").one()
            member_one = User(
                username="worktask_member_one",
                role="team_member",
                team_id=team_one.id,
                display_name="Project Member",
            )
            member_one.set_password("MemberPassword123!")
            member_two = User(
                username="worktask_member_two",
                role="team_member",
                team_id=team_two.id,
                display_name="Other Project Member",
            )
            member_two.set_password("MemberPassword123!")
            mentor = User(
                username="worktask_mentor",
                role="mentor",
                display_name="Project Mentor",
            )
            mentor.set_password("MentorPassword123!")
            db.session.add_all([member_one, member_two, mentor])
            db.session.flush()
            db.session.add(TeamMember(
                team_id=team_one.id,
                full_name="Project Member",
                prn="WORKTASKMEMBER001",
                email="project.member@example.edu",
                phone="+91 9876543210",
                user=member_one,
            ))
            db.session.add(MentorTeamAssignment(
                mentor_id=mentor.id, team_id=team_one.id
            ))
            db.session.commit()
            team_one_id = team_one.id
            member_one_id = member_one.id
            member_two_id = member_two.id

        mentor_client = self.app.test_client()
        self.assertEqual(mentor_client.post("/login", data={
            "username": "worktask_mentor",
            "password": "MentorPassword123!",
            "role": "mentor",
        }).status_code, 302)
        with patch.dict(self.app.config, {
            "TASK_EMAIL_SMTP_HOST": "smtp.example.edu",
            "TASK_EMAIL_SMTP_PORT": 587,
            "TASK_EMAIL_SMTP_USERNAME": "",
            "TASK_EMAIL_SMTP_PASSWORD": "",
            "TASK_EMAIL_FROM": "future-techclub@orchidengg.ac.in",
            "TASK_SMS_AUTH_KEY": "test-auth-key",
            "TASK_SMS_TEMPLATE_ID": "approved-template",
        }):
            with patch("app.routes.smtplib.SMTP") as smtp_factory:
                smtp_client = MagicMock()
                smtp_factory.return_value.__enter__.return_value = smtp_client
                sms_response = MagicMock()
                sms_response.__enter__.return_value = sms_response
                sms_response.read.return_value = b'{"data":{"type":"success"}}'
                with patch(
                    "app.routes.urllib.request.urlopen",
                    return_value=sms_response,
                ) as sms_request:
                    project_created = mentor_client.post(
                        f"/teams/{team_one_id}/work-tasks",
                        data={
                            "title": "Prepare project documentation",
                            "description": "Submit the updated code and setup notes.",
                            "due_date": "2027-01-15",
                            "assignee_ids": [str(member_one_id)],
                            "reference_file": (io.BytesIO(b"starter code"), "starter.py"),
                        },
                        content_type="multipart/form-data",
                    )
                smtp_factory.assert_called_once()
                smtp_client.starttls.assert_called_once()
                smtp_client.send_message.assert_called_once()
                sent_email = smtp_client.send_message.call_args.args[0]
                self.assertEqual(
                    sent_email["From"],
                    "future-techclub@orchidengg.ac.in",
                )
                self.assertEqual(sent_email["To"], "project.member@example.edu")
                self.assertEqual(
                    sent_email["Subject"],
                    "Task assigned: Prepare project documentation",
                )
                email_body = sent_email.get_content()
                self.assertIn("Submit the updated code and setup notes.", email_body)
                self.assertIn("15 Jan 2027", email_body)
                sms_request.assert_called_once()
                sms_request_object = sms_request.call_args.args[0]
                sms_payload = json.loads(sms_request_object.data)
                self.assertEqual(
                    sms_request_object.full_url,
                    "https://control.msg91.com/api/v5/flow",
                )
                self.assertEqual(
                    sms_request_object.get_header("Authkey"),
                    "test-auth-key",
                )
                self.assertEqual(
                    sms_request_object.get_header("Content-type"),
                    "application/json",
                )
                self.assertEqual(sms_payload["template_id"], "approved-template")
                self.assertEqual(sms_payload["recipients"], [{
                    "mobiles": "919876543210",
                    "VAR1": "Prepare project documentation",
                }])
        self.assertEqual(project_created.status_code, 302)
        workspace = mentor_client.get(f"/teams/{team_one_id}/workspace")
        self.assertIn(b"Prepare project documentation", workspace.data)

        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "worktask_member_one",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        dashboard_page = member_client.get("/member")
        self.assertIn(b"New task: Prepare project documentation", dashboard_page.data)
        self.assertIn(b"1 unread notification", dashboard_page.data)
        self.assertIn(b'aria-label="1 unread notification"', dashboard_page.data)
        task_page = member_client.get("/tasks")
        self.assertIn(b"Prepare project documentation", task_page.data)
        self.assertIn(b"New task: Prepare project documentation", task_page.data)
        with self.app.app_context():
            project_task = WorkTask.query.one()
            project_task_id = project_task.id
            assignment = WorkTaskAssignment.query.filter_by(
                task_id=project_task.id, user_id=member_one_id
            ).one()
            assignment_id = assignment.id
            self.assertEqual(project_task.reference_original_name, "starter.py")
            self.assertTrue(UserNotification.query.filter_by(
                user_id=member_one_id,
                title="New task: Prepare project documentation",
            ).first())
            self.assertFalse(UserNotification.query.filter_by(
                user_id=member_one_id,
                title="New task: Prepare project documentation",
                read=True,
            ).first())
            self.assertTrue(TeamChatMessage.query.filter(
                TeamChatMessage.team_id == team_one_id,
                TeamChatMessage.body.contains("Project task assigned"),
            ).first())
        self.assertEqual(mentor_client.post(
            f"/teams/{team_one_id}/work-tasks",
            data={
                "title": "Invalid cross-team task",
                "assignee_ids": [str(member_two_id)],
            },
        ).status_code, 400)
        reference = member_client.get(
            f"/work-tasks/{project_task_id}/files/reference"
        )
        self.assertEqual(reference.data, b"starter code")
        reference.close()

        submitted = member_client.post(
            f"/work-tasks/{project_task_id}/submit",
            data={"deliverable_file": (io.BytesIO(b"print('done')"), "solution.py")},
            content_type="multipart/form-data",
        )
        self.assertEqual(submitted.status_code, 302)
        with self.app.app_context():
            assignment = db.session.get(WorkTaskAssignment, assignment_id)
            self.assertEqual(assignment.status, "Submitted")
            self.assertEqual(assignment.deliverable_original_name, "solution.py")
        download = member_client.get(
            f"/work-tasks/{project_task_id}/files/deliverable/{assignment_id}"
        )
        self.assertEqual(download.data, b"print('done')")
        download.close()

        admin_client = self.app.test_client()
        self.assertEqual(admin_client.post("/login", data={
            "username": "admin",
            "password": "Admin@12345!",
            "role": "admin",
        }).status_code, 302)
        task_page_before_search = admin_client.get("/tasks")
        self.assertNotIn(b"worktask_member_two", task_page_before_search.data)
        self.assertNotIn(
            b"Type at least 2 characters to search active accounts.",
            task_page_before_search.data,
        )
        club_search = admin_client.get(
            "/tasks/search-users?q=worktask_member&scope=club"
        )
        self.assertEqual(club_search.status_code, 200)
        self.assertEqual(
            {user["id"] for user in club_search.json["users"]},
            {member_one_id, member_two_id},
        )
        project_search = admin_client.get(
            f"/tasks/search-users?q=worktask_member&scope=project&team_id={team_one_id}"
        )
        self.assertEqual(project_search.status_code, 200)
        self.assertEqual(
            {user["id"] for user in project_search.json["users"]},
            {member_one_id},
        )
        self.assertEqual(
            member_client.get("/tasks/search-users?q=worktask_member").status_code,
            403,
        )
        club_created = admin_client.post(
            "/tasks/create",
            data={
                "scope": "club",
                "title": "Club-wide preparation",
                "target_roles": ["team_member"],
            },
        )
        self.assertEqual(club_created.status_code, 302)
        admin_project_created = admin_client.post(
            "/tasks/create",
            data={
                "scope": "project",
                "team_id": str(team_one_id),
                "title": "Admin project review",
                "all_project_users": "1",
            },
        )
        self.assertEqual(admin_project_created.status_code, 302)
        admin_tasks = admin_client.get("/tasks")
        self.assertIn(b"Club-wide preparation", admin_tasks.data)
        self.assertIn(b"Admin project review", admin_tasks.data)
        with self.app.app_context():
            club_task = WorkTask.query.filter_by(scope="club").one()
            admin_project = WorkTask.query.filter_by(
                title="Admin project review"
            ).one()
            self.assertTrue(WorkTaskAssignment.query.filter_by(
                task_id=club_task.id, user_id=member_one_id
            ).first())
            self.assertTrue(WorkTaskAssignment.query.filter_by(
                task_id=club_task.id, user_id=member_two_id
            ).first())
            self.assertTrue(WorkTaskAssignment.query.filter_by(
                task_id=admin_project.id, user_id=member_one_id
            ).first())
        other_client = self.app.test_client()
        self.assertEqual(other_client.post("/login", data={
            "username": "worktask_member_two",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        self.assertEqual(other_client.get(
            f"/work-tasks/{project_task_id}/files/deliverable/{assignment_id}"
        ).status_code, 404)
        self.assertEqual(other_client.post(
            f"/teams/{team_one_id}/work-tasks",
            data={
                "title": "Cross-team task",
                "assignee_ids": [str(member_two_id)],
            },
        ).status_code, 404)

    def test_deadline_reminders_are_deduplicated_and_include_chat_and_email(self):
        today = datetime.today().date()
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            admin = User.query.filter_by(username="admin").one()
            member = User(
                username="deadline_member",
                role="team_member",
                team_id=team.id,
                email="deadline.member@example.edu",
            )
            member.set_password("DeadlinePassword123!")
            db.session.add(member)
            db.session.flush()
            db.session.add(TeamMember(
                team_id=team.id,
                full_name="Deadline Member",
                prn="DEADLINE001",
                user=member,
            ))
            for title, due_date, status in (
                ("Due today", today, "Open"),
                ("Due tomorrow", today + timedelta(days=1), "Open"),
                ("Already completed", today, "Completed"),
            ):
                task = WorkTask(
                    scope="project",
                    team_id=team.id,
                    assigner_id=admin.id,
                    title=title,
                    description=f"Instructions for {title.lower()}.",
                    due_date=due_date,
                )
                db.session.add(task)
                db.session.flush()
                db.session.add(WorkTaskAssignment(
                    task_id=task.id,
                    user_id=member.id,
                    status=status,
                ))
            db.session.add(TeamTask(
                team_id=team.id,
                assignee_id=member.id,
                assigner_id=admin.id,
                title="Team task deadline",
                description="Review the project plan.",
                due_date=today + timedelta(days=1),
            ))
            member_id = member.id
            db.session.commit()

        member_client = self.app.test_client()
        with patch.dict(self.app.config, {
            "TASK_EMAIL_SMTP_HOST": "smtp.example.edu",
            "TASK_EMAIL_SMTP_PORT": 587,
            "TASK_EMAIL_SMTP_USERNAME": "",
            "TASK_EMAIL_SMTP_PASSWORD": "",
            "TASK_EMAIL_FROM": "mailer@example.edu",
        }):
            with patch("app.routes.smtplib.SMTP") as smtp_factory:
                smtp_server = MagicMock()
                smtp_factory.return_value.__enter__.return_value = smtp_server
                self.assertEqual(member_client.post("/login", data={
                    "username": "deadline_member",
                    "password": "DeadlinePassword123!",
                    "role": "team_member",
                }).status_code, 302)
                first_page = member_client.get("/tasks")
                self.assertEqual(first_page.status_code, 200)
                self.assertIn(b"Due today", first_page.data)
                self.assertIn(b"Due tomorrow", first_page.data)
                self.assertIn(b"Already completed", first_page.data)
                self.assertEqual(smtp_server.send_message.call_count, 3)
                subjects = {
                    call.args[0]["Subject"]
                    for call in smtp_server.send_message.call_args_list
                }
                self.assertEqual(subjects, {
                    "Task deadline today: Due today",
                    "Task deadline tomorrow: Due tomorrow",
                    "Task deadline tomorrow: Team task deadline",
                })
                reminder_bodies = [
                    call.args[0].get_content()
                    for call in smtp_server.send_message.call_args_list
                ]
                self.assertTrue(any(
                    "Instructions for due today." in body
                    for body in reminder_bodies
                ))
                second_page = member_client.get("/tasks")
                self.assertEqual(second_page.status_code, 200)
                self.assertEqual(smtp_server.send_message.call_count, 3)
        with self.app.app_context():
            self.assertEqual(TaskDeadlineReminder.query.count(), 3)
            notifications = UserNotification.query.filter_by(
                user_id=member_id
            ).all()
            self.assertEqual(
                sum(notification.title.startswith("Task deadline") for notification in notifications),
                3,
            )
            group_messages = TeamChatMessage.query.filter(
                TeamChatMessage.body.like("%Automated task deadline reminder%")
            ).all()
            self.assertEqual(len(group_messages), 3)
            self.assertTrue(any("due today" in message.body for message in group_messages))
            self.assertTrue(any("due tomorrow" in message.body for message in group_messages))

    def test_task_assignment_roles_follow_admin_mentor_coordinator_leader_flow(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            team_two = Team.query.filter_by(team_code="FT_02").one()
            users = {}
            for key, role, assigned_team in (
                ("member", "team_member", team),
                ("other_member", "team_member", team_two),
                ("leader", "team_leader", team),
                ("mentor", "mentor", None),
                ("coordinator", "coordinator", None),
                ("judge", "judge", None),
            ):
                user = User(
                    username=f"taskflow_{key}",
                    role=role,
                    team_id=assigned_team.id if assigned_team else None,
                    display_name=f"Task Flow {key}",
                )
                user.set_password("TaskFlowPassword123!")
                db.session.add(user)
                users[key] = user
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(
                    mentor_id=users["mentor"].id,
                    team_id=team.id,
                ),
                CoordinatorTeamAssignment(
                    coordinator_id=users["coordinator"].id,
                    team_id=team.id,
                ),
            ])
            db.session.commit()
            team_id = team.id
            user_ids = {key: user.id for key, user in users.items()}

        def login_task_user(key, role):
            client = self.app.test_client()
            response = client.post("/login", data={
                "username": f"taskflow_{key}",
                "password": "TaskFlowPassword123!",
                "role": role,
            })
            self.assertEqual(response.status_code, 302)
            return client

        mentor_client = login_task_user("mentor", "mentor")
        mentor_workspace = mentor_client.get(f"/teams/{team_id}/workspace")
        for key in ("coordinator", "leader", "member"):
            self.assertIn(f"Task Flow {key}".encode(), mentor_workspace.data)
        self.assertNotIn(b"Task Flow mentor", mentor_workspace.data)
        self.assertNotIn(b"Task Flow other_member", mentor_workspace.data)
        mentor_all = mentor_client.post(
            f"/teams/{team_id}/work-tasks",
            data={"title": "Mentor assignment", "all_project_users": "1"},
        )
        self.assertEqual(mentor_all.status_code, 302)
        self.assertEqual(mentor_client.post(
            f"/teams/{team_id}/work-tasks",
            data={
                "title": "Mentor cannot assign to self",
                "assignee_ids": [str(user_ids["mentor"])],
            },
        ).status_code, 400)

        coordinator_client = login_task_user("coordinator", "coordinator")
        coordinator_all = coordinator_client.post(
            f"/teams/{team_id}/work-tasks",
            data={"title": "Coordinator assignment", "all_project_users": "1"},
        )
        self.assertEqual(coordinator_all.status_code, 302)
        self.assertEqual(coordinator_client.post(
            f"/teams/{team_id}/work-tasks",
            data={
                "title": "Coordinator cannot assign to mentor",
                "assignee_ids": [str(user_ids["mentor"])],
            },
        ).status_code, 400)

        leader_client = login_task_user("leader", "team_leader")
        leader_all = leader_client.post(
            f"/teams/{team_id}/work-tasks",
            data={"title": "Leader assignment", "all_project_users": "1"},
        )
        self.assertEqual(leader_all.status_code, 302)
        self.assertEqual(leader_client.post(
            f"/teams/{team_id}/work-tasks",
            data={
                "title": "Leader cannot assign to coordinator",
                "assignee_ids": [str(user_ids["coordinator"])],
            },
        ).status_code, 400)
        mentor_legacy = mentor_client.post(
            f"/teams/{team_id}/tasks",
            data={
                "title": "Mentor legacy assignment",
                "assignee_id": user_ids["coordinator"],
            },
        )
        self.assertEqual(mentor_legacy.status_code, 302)
        leader_legacy = mentor_client.post(
            f"/teams/{team_id}/tasks",
            data={
                "title": "Leader receives task",
                "assignee_id": user_ids["leader"],
            },
        )
        self.assertEqual(leader_legacy.status_code, 302)
        coordinator_legacy = coordinator_client.post(
            f"/teams/{team_id}/tasks",
            data={
                "title": "Coordinator legacy assignment",
                "assignee_id": user_ids["leader"],
            },
        )
        self.assertEqual(coordinator_legacy.status_code, 302)
        leader_to_member = leader_client.post(
            f"/teams/{team_id}/tasks",
            data={
                "title": "Leader legacy assignment",
                "assignee_id": user_ids["member"],
            },
        )
        self.assertEqual(leader_to_member.status_code, 302)
        self.assertEqual(leader_client.post(
            f"/teams/{team_id}/tasks",
            data={
                "title": "Leader cannot assign coordinator",
                "assignee_id": user_ids["coordinator"],
            },
        ).status_code, 404)
        with self.app.app_context():
            leader_task_id = TeamTask.query.filter_by(
                title="Leader receives task"
            ).one().id
        self.assertEqual(leader_client.post(
            f"/tasks/{leader_task_id}/status",
            data={"status": "Completed"},
        ).status_code, 302)

        admin_client = self.app.test_client()
        self.assertEqual(admin_client.post("/login", data={
            "username": "admin",
            "password": "Admin@12345!",
            "role": "admin",
        }).status_code, 302)
        admin_task = admin_client.post(
            "/tasks/create",
            data={
                "scope": "club",
                "title": "Admin eligible roles",
                "target_roles": [
                    "admin", "mentor", "coordinator", "team_leader",
                    "team_member", "judge",
                ],
            },
        )
        self.assertEqual(admin_task.status_code, 400)
        admin_task = admin_client.post(
            "/tasks/create",
            data={
                "scope": "club",
                "title": "Admin eligible roles",
                "target_roles": [
                    "admin", "mentor", "coordinator", "team_leader",
                    "team_member",
                ],
            },
        )
        self.assertEqual(admin_task.status_code, 302)

        with self.app.app_context():
            task_assignee_roles = {
                task.title: {
                    db.session.get(User, assignment.user_id).role
                    for assignment in task.assignments
                }
                for task in WorkTask.query.filter(
                    WorkTask.title.in_([
                        "Mentor assignment",
                        "Coordinator assignment",
                        "Leader assignment",
                    ])
                ).all()
            }
            self.assertEqual(
                task_assignee_roles["Mentor assignment"],
                {"coordinator", "team_leader", "team_member"},
            )
            self.assertEqual(
                task_assignee_roles["Coordinator assignment"],
                {"team_leader", "team_member"},
            )
            self.assertEqual(
                task_assignee_roles["Leader assignment"],
                {"team_member"},
            )
            admin_assignments = WorkTaskAssignment.query.join(
                WorkTask, WorkTask.id == WorkTaskAssignment.task_id
            ).filter(WorkTask.title == "Admin eligible roles").all()
            admin_assigned_roles = {
                db.session.get(User, assignment.user_id).role
                for assignment in admin_assignments
            }
            self.assertEqual(
                admin_assigned_roles,
                {"admin", "mentor", "coordinator", "team_leader", "team_member"},
            )

    def test_team_chat_supports_any_file_live_polling_and_five_minute_edits(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            member = User(
                username="chat_file_member",
                role="team_member",
                team_id=team.id,
            )
            member.set_password("MemberPassword123!")
            db.session.add(TeamMember(
                team_id=team.id,
                full_name="Chat File Member",
                prn="CHATFILE001",
                user=member,
            ))
            db.session.commit()
            team_id = team.id

        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        response = self.client.post(
            f"/teams/{team_id}/chat",
            data={
                "body": "See this file",
                "attachment": (io.BytesIO(b"arbitrary file contents"), "notes.anything"),
            },
            headers={"Accept": "application/json"},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200)
        message = response.json["message"]
        self.assertEqual(message["attachment_name"], "notes.anything")
        self.assertTrue(message["can_edit"])
        message_id = message["id"]
        poll = self.client.get(f"/teams/{team_id}/chat/messages")
        self.assertEqual(poll.status_code, 200)
        self.assertEqual(poll.json["messages"][0]["body"], "See this file")
        download = self.client.get(message["attachment_url"])
        self.assertEqual(download.status_code, 200)
        self.assertIn(b"attachment", download.headers.get("Content-Disposition", "").lower().encode())
        self.assertEqual(download.data, b"arbitrary file contents")
        download.close()

        edited = self.client.post(
            f"/team-chat/{message_id}/edit",
            data={"body": "Updated text"},
            headers={"Accept": "application/json"},
        )
        self.assertEqual(edited.status_code, 200)
        self.assertEqual(edited.json["message"]["body"], "Updated text")
        self.assertTrue(edited.json["message"]["edited"])
        with self.app.app_context():
            stored_message = db.session.get(TeamChatMessage, message_id)
            stored_message.sent_at = datetime.utcnow() - timedelta(minutes=6)
            db.session.commit()
        expired_edit = self.client.post(
            f"/team-chat/{message_id}/edit",
            data={"body": "Too late"},
        )
        self.assertEqual(expired_edit.status_code, 409)
        with patch("app.routes.CHAT_FILE_MAX_BYTES", 4):
            oversized = self.client.post(
                f"/teams/{team_id}/chat",
                data={
                    "attachment": (io.BytesIO(b"12345"), "too-big.bin"),
                },
                headers={"Accept": "application/json"},
                content_type="multipart/form-data",
            )
        self.assertEqual(oversized.status_code, 413)

        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "chat_file_member",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        member_download = member_client.get(message["attachment_url"])
        self.assertEqual(member_download.status_code, 200)
        member_download.close()
        with open(
            os.path.join(
                os.path.dirname(__file__),
                "..",
                "app",
                "static",
                "js",
                "chat.js",
            ),
            encoding="utf-8",
        ) as chat_script:
            self.assertIn("setInterval(refreshMessages, 1000)", chat_script.read())

    def test_admin_can_impersonate_accounts_in_view_only_and_full_audited_modes(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            target = User(
                username="impersonation_member",
                role="team_member",
                display_name="Impersonation Member",
                team_id=team.id,
            )
            target.set_password("MemberPassword123!")
            db.session.add(target)
            db.session.commit()
            target_id = target.id
            admin_id = User.query.filter_by(username="admin").one().id

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        dashboard_page = self.client.get("/admin")
        self.assertIn(b'href="/admin/certificates"', dashboard_page.data)
        self.assertIn(b"Certificate Management", dashboard_page.data)
        self.assertIn(b">Certificates</a>", dashboard_page.data)
        self.assertNotIn(b"All user accounts", dashboard_page.data)
        self.assertNotIn(b"Accounts and credentials", dashboard_page.data)
        self.assertIn(b"User Access", dashboard_page.data)
        self.assertIn(b'href="/admin/users#user-access"', dashboard_page.data)
        self.assertIn(b"/admin/users#user-access", dashboard_page.data)
        self.assertIn(b'class="collapse navbar-collapse" id="siteNavigation"', dashboard_page.data)
        self.assertIn(b'data-bs-target="#siteNavigation"', dashboard_page.data)
        self.assertNotIn(b"navbar-expand-lg", dashboard_page.data)
        self.assertGreater(
            dashboard_page.data.index(b"User Access"),
            dashboard_page.data.index(b'id="siteNavigation"'),
        )
        account_page = self.client.get("/admin/users")
        self.assertIn(b"View-only", account_page.data)
        self.assertIn(b"Full access", account_page.data)
        self.assertIn(b"No fixed time limit", account_page.data)
        self.assertNotIn(b"30 minutes", account_page.data)

        started = self.client.post(
            f"/admin/users/{target_id}/impersonate",
            data={"mode": "view_only"},
        )
        self.assertEqual(started.status_code, 302)
        with self.client.session_transaction() as browser_session:
            browser_session["admin_impersonation_expires_at"] = 1
        member_dashboard = self.client.get("/member")
        self.assertEqual(member_dashboard.status_code, 200)
        self.assertIn(b"Admin impersonation", member_dashboard.data)
        self.assertIn(b"no fixed time limit", member_dashboard.data)
        self.assertEqual(
            self.client.post(
                "/account/profile",
                data={
                    "display_name": "Changed in read-only mode",
                    "email": "changed@example.edu",
                    "phone": "",
                },
            ).status_code,
            403,
        )
        self.assertEqual(
            self.client.post("/admin/impersonation/upgrade").status_code,
            302,
        )
        upgraded_dashboard = self.client.get("/member")
        self.assertEqual(upgraded_dashboard.status_code, 200)
        self.assertIn(b"full access", upgraded_dashboard.data)
        self.assertIn(b"Actions are audit-logged", upgraded_dashboard.data)
        self.assertEqual(
            self.client.post(
                "/account/profile",
                data={
                    "display_name": "Updated after switching access",
                    "email": "switched@example.edu",
                    "phone": "",
                },
            ).status_code,
            302,
        )
        with self.app.app_context():
            switched_target = db.session.get(User, target_id)
            self.assertEqual(
                switched_target.display_name,
                "Updated after switching access",
            )
            self.assertTrue(AuditLog.query.filter_by(
                user_id=admin_id,
                action="Upgraded admin impersonation to full access",
            ).first())
        self.assertEqual(
            self.client.post("/admin/impersonation/end").status_code,
            302,
        )
        self.assertEqual(self.client.get("/admin").status_code, 200)

        self.assertEqual(
            self.client.post(
                f"/admin/users/{target_id}/impersonate",
                data={"mode": "full_access"},
            ).status_code,
            302,
        )
        self.assertEqual(
            self.client.post(
                "/account/profile",
                data={
                    "display_name": "Updated by administrator",
                    "email": "updated@example.edu",
                    "phone": "+919876543210",
                },
            ).status_code,
            302,
        )
        with self.app.app_context():
            updated_target = db.session.get(User, target_id)
            self.assertEqual(updated_target.display_name, "Updated by administrator")
            self.assertTrue(AuditLog.query.filter(
                AuditLog.user_id == admin_id,
                AuditLog.action == "Impersonated POST request",
                AuditLog.details.contains("impersonation_member"),
            ).first())
        self.assertEqual(
            self.client.post("/admin/impersonation/end").status_code,
            302,
        )
        self.assertEqual(self.client.get("/admin").status_code, 200)

    def test_team_project_details_keep_rich_text_and_remove_unsafe_html(self):
        project_details = (
            '<p style="text-align: center; position: fixed; color: rgb(1, 2, 3)">'
            "Reduce water waste</p><ul><li>Monitor usage</li>"
            "<li><strong>Detect leaks</strong></li></ul>"
            '<font size="5" color="#ff0000">Important</font>'
            '<a href="https://example.com" onclick="alert(1)">Project source</a>'
            '<a href="javascript:alert(1)">Unsafe link</a>'
            '<table><tbody><tr><th>Technology</th><td colspan="2">Python</td></tr></tbody></table>'
            "<script>alert('unsafe')</script>"
            "<img src=x onerror=alert('unsafe')>"
        )
        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        response = self.client.post(
            "/team/update",
            data={
                "name": "FT_01",
                "problem_statement": project_details,
            },
        )
        self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            self.assertEqual(
                team.problem_statement,
                '<p style="text-align: center; color: rgb(1, 2, 3)">'
                "Reduce water waste</p><ul><li>Monitor usage</li>"
                "<li><strong>Detect leaks</strong></li></ul>"
                '<span style="color: #ff0000; font-size: 18pt">Important</span>'
                '<a href="https://example.com">Project source</a><a>Unsafe link</a>'
                '<table><tbody><tr><th>Technology</th><td colspan="2">Python</td></tr></tbody></table>',
            )

        team_page = self.client.get("/team")
        self.assertIn(b"Bulleted list", team_page.data)
        self.assertIn(b"Text color", team_page.data)
        self.assertIn(b"Highlight color", team_page.data)
        self.assertIn(b"Insert table", team_page.data)
        self.assertIn(b"Merge cells", team_page.data)
        self.assertIn(b"Table rows", team_page.data)
        self.assertIn(b"Table columns", team_page.data)
        self.assertIn(b"Insert link", team_page.data)
        self.assertIn(b"Undo", team_page.data)
        self.assertIn(b"Indent", team_page.data)
        self.assertIn(b"Clear formatting", team_page.data)
        self.assertIn(
            b'<p style="text-align: center; color: rgb(1, 2, 3)">'
            b"Reduce water waste</p><ul><li>Monitor usage</li>"
            b"<li><strong>Detect leaks</strong></li></ul>"
            b'<span style="color: #ff0000; font-size: 18pt">Important</span>'
            b'<a href="https://example.com">Project source</a><a>Unsafe link</a>'
            b'<table><tbody><tr><th>Technology</th><td colspan="2">Python</td></tr></tbody></table>',
            team_page.data,
        )
        self.assertNotIn(b"<script>", team_page.data)
        self.assertNotIn(b"onerror", team_page.data)
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        admin_team_page = self.client.get("/admin/teams")
        self.assertEqual(admin_team_page.status_code, 200)
        self.assertIn(b"data-rich-editor", admin_team_page.data)
        self.assertIn(b"Numbered list", admin_team_page.data)

    def test_admin_can_search_any_account_and_users_can_reply_with_private_files(self):
        with self.app.app_context():
            target = User(
                username="searchable_judge",
                role="judge",
                display_name="Searchable Judge",
                email="judge@example.edu",
            )
            target.set_password("JudgePassword123!")
            unrelated = User(
                username="unrelated_member",
                role="team_member",
                display_name="Unrelated Member",
            )
            unrelated.set_password("MemberPassword123!")
            db.session.add_all([target, unrelated])
            db.session.commit()
            target_id = target.id
            unrelated_id = unrelated.id
            admin_id = User.query.filter_by(role="admin").first().id

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        search_page = self.client.get("/admin/chats?q=Searchable%20Judge")
        self.assertEqual(search_page.status_code, 200)
        self.assertIn(b"Searchable Judge", search_page.data)
        self.assertNotIn(b"Unrelated Member", search_page.data)
        judge_filter = self.client.get("/admin/chats?role=judge")
        self.assertIn(b"Searchable Judge", judge_filter.data)
        self.assertNotIn(b"Unrelated Member", judge_filter.data)
        member_filter = self.client.get("/admin/chats?role=team_member")
        self.assertIn(b"Unrelated Member", member_filter.data)
        self.assertNotIn(b"Searchable Judge", member_filter.data)
        self.assertEqual(self.client.get("/admin/chats?role=invalid").status_code, 400)
        selected_page = self.client.get(f"/admin/chats?user_id={target_id}")
        self.assertEqual(selected_page.status_code, 200)
        self.assertIn(b"Font size", selected_page.data)
        self.assertIn(b"Increase message font size", selected_page.data)
        self.assertNotIn(b"Any file type", selected_page.data)
        sent = self.client.post(
            f"/admin/chats/{target_id}/send",
            data={
                "body": "Admin private note",
                "attachment": (io.BytesIO(b"private admin attachment"), "brief.bin"),
            },
            headers={"Accept": "application/json"},
            content_type="multipart/form-data",
        )
        self.assertEqual(sent.status_code, 200)
        admin_message = sent.json["message"]
        self.assertEqual(admin_message["attachment_name"], "brief.bin")
        admin_message_id = admin_message["id"]

        target_client = self.app.test_client()
        self.assertEqual(target_client.post("/login", data={
            "username": "searchable_judge",
            "password": "JudgePassword123!",
            "role": "judge",
        }).status_code, 302)
        inbox = target_client.get(f"/chats?admin_id={admin_id}")
        self.assertEqual(inbox.status_code, 200)
        target_history = target_client.get("/account/history")
        self.assertEqual(target_history.status_code, 200)
        self.assertIn(b"New chat message", target_history.data)
        self.assertIn(b"admin", target_history.data.lower())
        self.assertNotIn(b"Admin private note", target_history.data)
        self.assertEqual(
            target_client.get(f"/account/history?user_id={unrelated_id}").status_code,
            403,
        )
        messages = target_client.get(f"/chats/{admin_id}/messages")
        self.assertEqual(messages.json["messages"][0]["body"], "Admin private note")
        admin_attachment = target_client.get(admin_message["attachment_url"])
        self.assertEqual(admin_attachment.data, b"private admin attachment")
        admin_attachment.close()
        reply = target_client.post(
            f"/chats/{admin_id}/send",
            data={"body": "I received it"},
            headers={"Accept": "application/json"},
        )
        self.assertEqual(reply.status_code, 200)
        admin_messages = self.client.get(f"/admin/chats/{target_id}/messages")
        self.assertEqual(
            [message["body"] for message in admin_messages.json["messages"]],
            ["Admin private note", "I received it"],
        )
        admin_history = self.client.get("/account/history")
        self.assertEqual(admin_history.status_code, 200)
        self.assertIn(b"User activity history", admin_history.data)
        self.assertIn(b"searchable_judge", admin_history.data)
        selected_history = self.client.get(
            f"/account/history?user_id={target_id}"
        )
        self.assertEqual(selected_history.status_code, 200)
        self.assertIn(b"New chat message", selected_history.data)
        self.assertNotIn(b"Admin private note", selected_history.data)

        unrelated_client = self.app.test_client()
        self.assertEqual(unrelated_client.post("/login", data={
            "username": "unrelated_member",
            "password": "MemberPassword123!",
            "role": "team_member",
        }).status_code, 302)
        self.assertEqual(
            unrelated_client.get(admin_message["attachment_url"]).status_code, 404
        )
        self.assertEqual(
            unrelated_client.post(
                f"/admin/chat-messages/{admin_message_id}/edit",
                data={"body": "Unauthorized"},
            ).status_code,
            403,
        )
        with self.app.app_context():
            message = db.session.get(AdminChatMessage, admin_message_id)
            message.sent_at = datetime.utcnow() - timedelta(minutes=6)
            db.session.commit()
        self.assertEqual(
            self.client.post(
                f"/admin/chat-messages/{admin_message_id}/edit",
                data={"body": "Too late"},
            ).status_code,
            409,
        )
        with self.app.app_context():
            self.assertNotEqual(
                db.session.get(User, target_id).id,
                db.session.get(User, unrelated_id).id,
            )

    def test_role_based_direct_chat_contacts_are_scoped_and_enforced(self):
        with self.app.app_context():
            team_one = Team.query.filter_by(team_code="FT_01").one()
            team_two = Team.query.filter_by(team_code="FT_02").one()
            accounts = {}
            for key, role, team in (
                ("mentor_a", "mentor", None),
                ("mentor_b", "mentor", None),
                ("coordinator_a", "coordinator", None),
                ("coordinator_b", "coordinator", None),
                ("leader_a", "team_leader", team_one),
                ("leader_b", "team_leader", team_two),
                ("member_a", "team_member", team_one),
                ("member_b", "team_member", team_two),
                ("judge", "judge", None),
            ):
                user = User(
                    username=f"rolechat_{key}",
                    role=role,
                    team_id=team.id if team else None,
                    display_name=f"RoleChat {key}",
                )
                user.set_password("RoleChatPassword123!")
                db.session.add(user)
                accounts[key] = user
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(
                    mentor_id=accounts["mentor_a"].id,
                    team_id=team_one.id,
                ),
                MentorTeamAssignment(
                    mentor_id=accounts["mentor_b"].id,
                    team_id=team_two.id,
                ),
                CoordinatorTeamAssignment(
                    coordinator_id=accounts["coordinator_a"].id,
                    team_id=team_one.id,
                ),
                CoordinatorTeamAssignment(
                    coordinator_id=accounts["coordinator_b"].id,
                    team_id=team_two.id,
                ),
            ])
            db.session.commit()
            account_ids = {key: user.id for key, user in accounts.items()}
            account_names = {
                key: f"RoleChat {key}".encode()
                for key in accounts
            }

        def signed_in_account(key, role):
            client = self.app.test_client()
            response = client.post("/login", data={
                "username": f"rolechat_{key}",
                "password": "RoleChatPassword123!",
                "role": role,
            })
            self.assertEqual(response.status_code, 302)
            return client

        role_expectations = {
            "mentor_a": (
                "mentor",
                {"mentor_b", "coordinator_a", "coordinator_b", "leader_a", "leader_b", "member_a"},
                {"member_b", "judge"},
            ),
            "coordinator_a": (
                "coordinator",
                {"mentor_a", "mentor_b", "coordinator_b", "leader_a", "member_a"},
                {"leader_b", "member_b", "judge"},
            ),
            "leader_a": (
                "team_leader",
                {"mentor_a", "mentor_b", "coordinator_a", "member_a"},
                {"coordinator_b", "leader_b", "member_b", "judge"},
            ),
            "member_a": (
                "team_member",
                {"mentor_a", "mentor_b", "coordinator_a", "leader_a"},
                {"coordinator_b", "leader_b", "member_b", "judge"},
            ),
            "judge": (
                "judge",
                set(),
                {"mentor_a", "coordinator_a", "leader_a", "member_a"},
            ),
        }
        for key, (role, expected, excluded) in role_expectations.items():
            with self.subTest(role=role):
                client = signed_in_account(key, role)
                dashboard = client.get("/dashboard", follow_redirects=True)
                self.assertEqual(dashboard.status_code, 200)
                inbox = client.get("/chats")
                self.assertEqual(inbox.status_code, 200)
                for contact in expected:
                    self.assertIn(account_names[contact], inbox.data)
                for contact in excluded:
                    self.assertNotIn(account_names[contact], inbox.data)
                denied_contact = excluded.pop() if excluded else None
                if denied_contact:
                    self.assertEqual(
                        client.get(
                            f"/chats?user_id={account_ids[denied_contact]}"
                        ).status_code,
                        404,
                    )

        mentor_client = signed_in_account("mentor_a", "mentor")
        allowed_contact_id = account_ids["member_a"]
        allowed_page = mentor_client.get(f"/chats?user_id={allowed_contact_id}")
        self.assertEqual(allowed_page.status_code, 200)
        allowed_send = mentor_client.post(
            f"/chats/{allowed_contact_id}/send",
            data={"body": "Mentor to assigned member"},
            headers={"Accept": "application/json"},
        )
        self.assertEqual(allowed_send.status_code, 200)
        self.assertEqual(
            mentor_client.get(
                f"/chats/{account_ids['member_b']}/messages"
            ).status_code,
            404,
        )

        admin_client = self.app.test_client()
        self.assertEqual(admin_client.post("/login", data={
            "username": "admin",
            "password": "Admin@12345!",
            "role": "admin",
        }).status_code, 302)
        self.assertEqual(
            admin_client.get("/dashboard", follow_redirects=True).status_code,
            200,
        )
        admin_chat_page = admin_client.get("/admin/chats")
        for contact in account_names:
            self.assertIn(account_names[contact], admin_chat_page.data)

    def test_seed_migrates_legacy_admin_account_to_configured_username(self):
        with self.app.app_context(), patch.dict(
            self.app.config,
            {
                "BOOTSTRAP_ADMIN_USERNAME": "24062231242187",
                "BOOTSTRAP_ADMIN_PASSWORD": "@24062231242187",
            },
        ):
            seed_initial_data()
            admin = User.query.filter_by(username="24062231242187").one()
            self.assertEqual(admin.role, "admin")
            self.assertTrue(admin.check_password("@24062231242187"))
            self.assertIsNone(User.query.filter_by(username="admin").first())

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 200)
        self.assertEqual(self.login("24062231242187", "@24062231242187").status_code, 302)

    def test_import_student_roster_creates_team_scoped_student_logins(self):
        roster = load_roster()
        self.assertEqual(len(roster), 130)
        self.assertEqual(import_roster(roster, self.app), 130)
        with self.app.app_context():
            member = User.query.filter_by(username="SM26_1_033").one()
            member.set_password("ChangedMemberPassword123!")
            db.session.commit()
        self.assertEqual(import_roster(roster, self.app), 130)
        with self.app.app_context():
            seed_initial_data()

        with self.app.app_context():
            self.assertEqual(TeamMember.query.count(), 130)
            self.assertEqual(User.query.filter_by(role="team_leader").count(), 27)
            self.assertEqual(User.query.filter_by(role="team_member").count(), 103)
            self.assertIsNone(User.query.filter_by(username="FT_01").first())
            leader = User.query.filter_by(username="SM26_1_094").one()
            member = User.query.filter_by(username="SM26_1_033").one()
            team = Team.query.filter_by(team_code="FT_01").one()
            self.assertEqual(leader.role, "team_leader")
            self.assertEqual(leader.team_id, team.id)
            self.assertTrue(leader.check_password("FT_01"))
            self.assertEqual(member.role, "team_member")
            self.assertEqual(member.team_id, team.id)
            self.assertTrue(member.check_password("ChangedMemberPassword123!"))
            self.assertEqual(
                User.query.filter_by(role="team_leader").count(),
                27,
            )
            self.assertEqual(team.domain.code, "D05")
            self.assertEqual(team.leader_name, "Navgire Pratiksha")

    def create_judge_assignment(self, team_code="FT_01"):
        with self.app.app_context():
            judge = User(
                username="judge_test",
                role="judge",
                display_name="Test Judge",
            )
            judge.set_password("JudgePassword123!")
            team = Team.query.filter_by(team_code=team_code).one()
            db.session.add(judge)
            db.session.flush()
            db.session.add(JudgeAssignment(judge_id=judge.id, team_id=team.id))
            db.session.commit()
            return team.id

    def test_judges_can_only_evaluate_assigned_teams_and_finalized_scores_lock(self):
        assigned_team_id = self.create_judge_assignment()
        with self.app.app_context():
            unassigned_team_id = Team.query.filter_by(team_code="FT_02").one().id
        self.assertEqual(self.login("judge_test", "JudgePassword123!").status_code, 302)

        dashboard = self.client.get("/judge")
        self.assertIn(b"FT_01", dashboard.data)
        self.assertNotIn(b"FT_02", dashboard.data)
        self.assertEqual(self.client.get(f"/judge/evaluate/{unassigned_team_id}").status_code, 404)

        marks = {
            "problem": 8,
            "innovation": 10,
            "implementation": 12,
            "prototype": 15,
            "impact": 5,
            "presentation": 5,
            "qa": 5,
        }
        self.assertEqual(
            self.client.post(f"/judge/evaluate/{assigned_team_id}", data=marks).status_code,
            302,
        )
        with self.app.app_context():
            evaluation = Evaluation.query.one()
            self.assertTrue(evaluation.finalized)
            self.assertFalse(evaluation.approved)
        self.assertEqual(
            self.client.post(f"/judge/evaluate/{assigned_team_id}", data=marks).status_code,
            409,
        )
        dashboard = self.client.get("/judge")
        self.assertIn(b"1</h2><div class=\"stat-label\">Awaiting admin approval", dashboard.data)
        self.assertIn(b"0</h2><div class=\"stat-label\">Remaining", dashboard.data)
        self.assertIn(b"60.0 / 60", dashboard.data)
        score_page = self.client.get(f"/judge/evaluate/{assigned_team_id}")
        self.assertIn(b"waiting for admin approval", score_page.data)
        self.assertIn(b"Submitted marks", score_page.data)
        self.assertIn(b'<td class="text-end">8.0</td>', score_page.data)
        self.assertEqual(
            self.client.post(
                f"/judge/evaluate/{assigned_team_id}/reopen",
                data={"csrf_token": "test-token"},
            ).status_code,
            302,
        )
        reopened_page = self.client.get(f"/judge/evaluate/{assigned_team_id}")
        self.assertIn(b"Update the marks below and submit again", reopened_page.data)
        updated_marks = dict(marks, problem=7)
        self.assertEqual(
            self.client.post(
                f"/judge/evaluate/{assigned_team_id}",
                data=updated_marks,
            ).status_code,
            302,
        )
        with self.app.app_context():
            evaluation = Evaluation.query.one()
            self.assertEqual(evaluation.total, 59)
            self.assertEqual(evaluation.problem, 7)
            self.assertTrue(evaluation.finalized)
            self.assertFalse(evaluation.approved)
            evaluation_id = evaluation.id

        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        admin_dashboard = self.client.get("/admin")
        self.assertIn(b"Judge marks awaiting approval", admin_dashboard.data)
        self.assertIn(b"59.0 / 60", admin_dashboard.data)
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{evaluation_id}/approve",
                data={"csrf_token": "test-token"},
            ).status_code,
            302,
        )
        with self.app.app_context():
            evaluation = Evaluation.query.one()
            self.assertTrue(evaluation.approved)
            self.assertEqual(
                evaluation.approver.username,
                "admin",
            )

        with self.app.app_context():
            additional_team = Team.query.filter_by(team_code="FT_02").one()
            additional_team.domain_id = Domain.query.order_by(Domain.id).first().id
            judge_id = User.query.filter_by(username="judge_test").one().id
            db.session.add(
                JudgeAssignment(judge_id=judge_id, team_id=additional_team.id)
            )
            db.session.add(
                Evaluation(
                    team_id=additional_team.id,
                    judge_id=judge_id,
                    total=50,
                    finalized=True,
                )
            )
            db.session.commit()
            returned_evaluation_id = Evaluation.query.filter_by(
                team_id=additional_team.id
            ).one().id
            returned_team_id = additional_team.id
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{returned_evaluation_id}/return",
                data={"csrf_token": "test-token"},
            ).status_code,
            302,
        )
        with self.app.app_context():
            returned_evaluation = db.session.get(Evaluation, returned_evaluation_id)
            self.assertFalse(returned_evaluation.finalized)
            self.assertFalse(returned_evaluation.approved)

        self.client.post("/logout")
        self.assertEqual(self.login("judge_test", "JudgePassword123!").status_code, 302)
        self.assertEqual(
            self.client.post(
                f"/judge/evaluate/{assigned_team_id}/reopen",
                data={"csrf_token": "test-token"},
            ).status_code,
            409,
        )
        approved_page = self.client.get(f"/judge/evaluate/{assigned_team_id}")
        self.assertIn(b"Evaluation submitted.", approved_page.data)
        self.assertEqual(self.client.get(f"/judge/evaluate/{returned_team_id}").status_code, 200)

    def test_startup_does_not_treat_unattributed_judge_marks_as_approved(self):
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02"])
            ).order_by(Team.team_code).all()
            judge_one = User(username="approval_migration_judge_one", role="judge")
            judge_one.set_password("JudgePassword123!")
            judge_two = User(username="approval_migration_judge_two", role="judge")
            judge_two.set_password("JudgePassword123!")
            admin = User.query.filter_by(role="admin").one()
            db.session.add_all([judge_one, judge_two])
            db.session.flush()
            legacy_mark = Evaluation(
                team_id=teams[0].id,
                judge_id=judge_one.id,
                total=55,
                finalized=True,
                approved=True,
            )
            admin_approved_mark = Evaluation(
                team_id=teams[1].id,
                judge_id=judge_two.id,
                total=58,
                finalized=True,
                approved=True,
                approved_by=admin.id,
                approved_at=datetime.utcnow(),
            )
            db.session.add_all([legacy_mark, admin_approved_mark])
            db.session.commit()
            legacy_mark_id = legacy_mark.id
            admin_approved_mark_id = admin_approved_mark.id

            ensure_member_schema()
            db.session.expire_all()

            self.assertFalse(
                db.session.get(Evaluation, legacy_mark_id).approved
            )
            self.assertTrue(
                db.session.get(Evaluation, admin_approved_mark_id).approved
            )

    def test_judge_evaluation_must_be_approved_before_rankings_and_freeze_blocks_approval(self):
        team_id = self.create_judge_assignment()
        with self.app.app_context():
            team = db.session.get(Team, team_id)
            team.domain_id = Domain.query.order_by(Domain.id).first().id
            team.registration_status = "Approved"
            db.session.add(
                AdminEvaluation(team_id=team_id, total=40, finalized=True)
            )
            db.session.commit()
        self.assertEqual(self.login("judge_test", "JudgePassword123!").status_code, 302)
        marks = {
            "problem": 8,
            "innovation": 10,
            "implementation": 12,
            "prototype": 15,
            "impact": 5,
            "presentation": 5,
            "qa": 5,
        }
        self.client.post(f"/judge/evaluate/{team_id}", data=marks)
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.client.post("/admin/nominate")
        with self.app.app_context():
            self.assertEqual(Nomination.query.count(), 0)
        with self.app.app_context():
            evaluation_id = Evaluation.query.one().id
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{evaluation_id}/approve",
                data={"csrf_token": "test-token"},
            ).status_code,
            302,
        )
        self.client.post("/admin/nominate")
        with self.app.app_context():
            self.assertEqual(Nomination.query.count(), 1)
        self.client.post("/admin/publish-nominations")
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{evaluation_id}/approve",
                data={"csrf_token": "test-token"},
            ).status_code,
            409,
        )

    def test_admin_score_overview_shows_pending_inputs_and_adjustment_controls(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            team.domain_id = Domain.query.order_by(Domain.id).first().id
            judge = User(username="overview_judge", role="judge", display_name="Overview Judge")
            judge.set_password("JudgePassword123!")
            mentor = User(username="overview_mentor", role="mentor", display_name="Overview Mentor")
            mentor.set_password("MentorPassword123!")
            db.session.add_all([judge, mentor])
            db.session.flush()
            db.session.add_all([
                JudgeAssignment(judge_id=judge.id, team_id=team.id),
                MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id),
                Evaluation(
                    team_id=team.id,
                    judge_id=judge.id,
                    problem=7,
                    innovation=8,
                    implementation=9,
                    prototype=10,
                    impact=4,
                    presentation=3,
                    qa=2,
                    total=43,
                    finalized=True,
                    approved=False,
                ),
                MentorEvaluation(
                    team_id=team.id,
                    mentor_id=mentor.id,
                    development=9,
                    teamwork=8,
                    documentation=7,
                    discipline=6,
                    feasibility=5,
                    total=35,
                    status="Pending",
                ),
            ])
            team_id = team.id
            db.session.commit()
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        dashboard = self.client.get("/admin")
        for section in (
            "mentors", "coordinators", "teams", "members", "pending-teams",
            "judge-pending", "judge-approved", "admin-scores", "mentor-proposals",
            "videos", "winners", "scores",
        ):
            self.assertIn(f"/admin/overview/{section}".encode(), dashboard.data)
            self.assertEqual(
                self.client.get(f"/admin/overview/{section}").status_code,
                200,
                section,
            )
        judge_review_page = self.client.get("/admin/overview/judge-approved")
        self.assertIn(b"Judge evaluation marks &amp; approval", judge_review_page.data)
        self.assertIn(b"43.0", judge_review_page.data)
        self.assertIn(b"Awaiting approval", judge_review_page.data)
        self.assertIn(b"Approve marks", judge_review_page.data)
        self.assertIn(b"Edit marks", judge_review_page.data)
        self.assertIn(b"Mentor 40-mark proposals", dashboard.data)
        self.assertIn(b"Judge marks awaiting approval", dashboard.data)
        self.assertIn(b"Review / adjust", dashboard.data)
        self.assertIn(b"Approve marks", dashboard.data)
        score_page = self.client.get("/admin/overview/scores")
        self.assertEqual(score_page.status_code, 200)
        self.assertIn(b"Team_Id", score_page.data)
        self.assertIn(b"Leader", score_page.data)
        self.assertIn(b"40 marks", score_page.data)
        self.assertIn(b"60 marks", score_page.data)
        self.assertIn(b"100 marks", score_page.data)
        self.assertIn(b"35.0", score_page.data)
        self.assertIn(b"43.0", score_page.data)
        self.assertIn(b"78.0", score_page.data)
        with self.app.app_context():
            evaluation_id = Evaluation.query.one().id
            proposal_id = MentorEvaluation.query.one().id
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{evaluation_id}/edit",
                data={
                    "problem": "8",
                    "innovation": "8",
                    "implementation": "9",
                    "prototype": "10",
                    "impact": "4",
                    "presentation": "3",
                    "qa": "2",
                },
            ).status_code,
            302,
        )
        self.assertEqual(
            self.client.post(
                f"/admin/admin-evaluation/{team_id}",
                data={
                    "mentor_evaluation_id": str(proposal_id),
                    "development": "10",
                    "teamwork": "8",
                    "documentation": "7",
                    "discipline": "6",
                    "feasibility": "5",
                },
            ).status_code,
            302,
        )
        with self.app.app_context():
            self.assertTrue(Evaluation.query.one().approved)
            self.assertEqual(Evaluation.query.one().total, 44)
            self.assertEqual(AdminEvaluation.query.one().total, 36)
            self.assertEqual(MentorEvaluation.query.one().status, "Approved")
            evaluation_id = Evaluation.query.one().id
        admin_score_page = self.client.get("/admin/overview/admin-scores")
        self.assertIn(b"Admin marks", admin_score_page.data)
        self.assertIn(b"/ 40", admin_score_page.data)
        self.assertIn(b"36.0", admin_score_page.data)
        self.assertNotIn(b"60 marks", admin_score_page.data)
        self.assertNotIn(b"100 marks", admin_score_page.data)
        self.assertNotIn(b"43.0", admin_score_page.data)
        approved_page = self.client.get("/admin/overview/judge-approved")
        self.assertIn(b"Confirm approval", approved_page.data)
        self.assertIn(b"Edit marks", approved_page.data)
        self.assertEqual(
            self.client.post(
                f"/admin/judge-evaluations/{evaluation_id}/approve",
                data={"csrf_token": "test-token"},
            ).status_code,
            302,
        )
        with self.app.app_context():
            self.assertTrue(Evaluation.query.one().approved)

    def test_admin_can_cancel_all_published_results_and_return_to_score_review(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            domain = Domain.query.order_by(Domain.id).first()
            team.domain_id = domain.id
            admin = User.query.filter_by(role="admin").first()
            db.session.add_all([
                Nomination(team_id=team.id, rank=1, published=True),
                Winner(team_id=team.id, published=True),
                DomainResultPublication(domain_id=domain.id, published_by=admin.id),
                Announcement(
                    title=f"Domain Results Announced: {domain.code}",
                    body="Published scores",
                ),
                Announcement(
                    title="Final Winners & Leaderboard Announced",
                    body="Published winner",
                ),
                UserNotification(
                    user_id=admin.id,
                    title="Domain nominations published",
                    body="Published",
                ),
                UserNotification(
                    user_id=admin.id,
                    title=f"{domain.code} results announced",
                    body="Published",
                ),
                SiteSetting(key="nominations_published", value="true"),
                SiteSetting(key="nominations_generated", value="true"),
                SiteSetting(key="final_leaderboard_published", value="true"),
            ])
            db.session.commit()
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        response = self.client.post("/admin/cancel-published-results")
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.client.get("/admin").status_code, 200)
        with self.app.app_context():
            self.assertFalse(Nomination.query.one().published)
            self.assertEqual(Winner.query.count(), 0)
            self.assertEqual(DomainResultPublication.query.count(), 0)
            self.assertEqual(
                Announcement.query.filter(
                    Announcement.title.like("Domain Results Announced:%")
                ).count(),
                0,
            )
            self.assertEqual(
                db.session.get(SiteSetting, "nominations_published").value,
                "false",
            )
            self.assertEqual(
                db.session.get(SiteSetting, "nominations_generated").value,
                "false",
            )
            self.assertEqual(
                db.session.get(SiteSetting, "final_leaderboard_published").value,
                "false",
            )
        regenerated = self.client.post("/admin/nominate")
        self.assertEqual(regenerated.status_code, 302)

    def test_team_leader_is_limited_to_team_dashboard_and_admin_is_forbidden(self):
        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        response = self.client.get("/team")
        self.assertIn(b"FT_01", response.data)
        self.assertEqual(self.client.get("/admin").status_code, 403)

    def test_team_leader_can_register_member_and_get_one_time_login(self):
        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        response = self.client.post(
            "/team/members",
            data={
                "full_name": "Asha Student",
                "prn": "PRN/2026/001",
                "roll_number": "12",
                "department": "Computer Engineering",
                "year": "Third Year",
                "email": "asha@example.edu",
                "phone": "9000000000",
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Member registered", response.data)
        with self.app.app_context():
            member = TeamMember.query.filter_by(prn="PRN/2026/001").one()
            self.assertEqual(member.department, "Computer Engineering")
            self.assertEqual(member.user.role, "team_member")
            self.assertEqual(member.user.team_id, member.team_id)
            self.assertEqual(member.user.email, "asha@example.edu")
            self.assertEqual(member.user.phone, "9000000000")
            self.assertIn(member.user.username.encode(), response.data)
            self.assertIn(b"Password:", response.data)
            member_user_id = member.user_id
            member_record_id = member.id
        denied_contact_update = self.client.post(
            f"/admin/users/{member_user_id}/contact",
            data={"email": "unauthorized@example.edu", "phone": "9000000001"},
        )
        self.assertEqual(denied_contact_update.status_code, 403)
        admin_client = self.app.test_client()
        self.assertEqual(admin_client.post("/login", data={
            "username": "admin",
            "password": "Admin@12345!",
            "role": "admin",
        }).status_code, 302)
        admin_users_page = admin_client.get("/admin/users")
        self.assertEqual(admin_users_page.status_code, 200)
        self.assertIn(b"Edit contact", admin_users_page.data)
        contact_response = admin_client.post(
            f"/admin/users/{member_user_id}/contact",
            data={"email": "new-asha@example.edu", "phone": "+91 9000000000"},
        )
        self.assertEqual(contact_response.status_code, 302)
        with self.app.app_context():
            member = db.session.get(TeamMember, member_record_id)
            self.assertEqual(member.user.email, "new-asha@example.edu")
            self.assertEqual(member.user.phone, "+91 9000000000")
            self.assertEqual(member.email, "new-asha@example.edu")
            self.assertEqual(member.phone, "+91 9000000000")
        markup = response.data.decode()
        username = re.search(r"Username:\s*<code>([^<]+)", markup).group(1)
        password = re.search(r"Password:\s*<code>([^<]+)", markup).group(1)
        member_client = self.app.test_client()
        self.assertEqual(
            member_client.post(
                "/login",
                data={
                    "username": username,
                    "password": password,
                    "role": "team_member",
                },
            ).status_code,
            302,
        )
        self.assertEqual(member_client.get("/member").status_code, 200)

        duplicate = self.client.post(
            "/team/members",
            data={"full_name": "Another student", "prn": "prn/2026/001"},
        )
        self.assertEqual(duplicate.status_code, 302)
        with self.app.app_context():
            self.assertEqual(TeamMember.query.count(), 1)

    def test_admin_can_create_team_leader_and_member_for_selected_team(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            teams = Team.query.filter(Team.team_code.in_(["FT_01", "FT_02"])).order_by(Team.team_code).all()
            team_ids = {team.team_code: team.id for team in teams}

        leader_response = self.client.post("/admin/users", data={
            "username": "custom_ft_leader",
            "display_name": "Custom Team Leader",
            "role": "team_leader",
            "team_id": str(team_ids["FT_02"]),
            "password": "LeaderPass12",
        })
        self.assertEqual(leader_response.status_code, 302)
        member_response = self.client.post("/admin/users", data={
            "username": "custom_ft_member",
            "display_name": "Custom Team Member",
            "role": "team_member",
            "team_id": str(team_ids["FT_02"]),
            "full_name": "Custom Team Member",
            "prn": "CUSTOM002",
            "password": "MemberPass1!",
        })
        self.assertEqual(member_response.status_code, 302)

        with self.app.app_context():
            leader = User.query.filter_by(username="custom_ft_leader").one()
            member_account = User.query.filter_by(username="custom_ft_member").one()
            member = TeamMember.query.filter_by(user_id=member_account.id).one()
            self.assertEqual(leader.team_id, team_ids["FT_02"])
            self.assertEqual(member_account.team_id, team_ids["FT_02"])
            self.assertEqual(member.team_id, team_ids["FT_02"])
            self.assertEqual(member.prn, "CUSTOM002")

    def test_existing_username_updates_only_when_both_role_and_password_change(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02", "FT_03"])
            ).all()
            team_ids = {team.team_code: team.id for team in teams}
        initial = {
            "username": "role_change_user",
            "display_name": "Role Change User",
            "role": "mentor",
            "team_ids": [str(team_ids["FT_01"]), str(team_ids["FT_02"])],
            "password": "InitPass12",
        }
        self.assertEqual(self.client.post("/admin/users", data=initial).status_code, 302)

        same_role = dict(initial, password="Changed1234")
        self.client.post("/admin/users", data=same_role)
        same_password = dict(
            initial,
            role="coordinator",
            team_ids=[str(team_ids["FT_03"])],
        )
        self.client.post("/admin/users", data=same_password)
        with self.app.app_context():
            account = User.query.filter_by(username="role_change_user").one()
            self.assertEqual(account.role, "mentor")
            self.assertTrue(account.check_password("InitPass12"))

        updated = dict(
            initial,
            role="coordinator",
            team_ids=[str(team_ids["FT_03"])],
            password="NewRole1234",
        )
        self.client.post("/admin/users", data=updated)
        with self.app.app_context():
            account = User.query.filter_by(username="role_change_user").one()
            self.assertEqual(account.role, "coordinator")
            self.assertTrue(account.check_password("NewRole1234"))
            self.assertEqual(
                {assignment.team_id for assignment in account.coordinator_team_assignments},
                {team_ids["FT_03"]},
            )
            self.assertEqual(account.mentor_team_assignments, [])

    def test_admin_dashboard_shows_mentor_coordinator_counts_status_and_teams(self):
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02"])
            ).order_by(Team.team_code).all()
            mentor = User(
                username="dashboard_mentor",
                role="mentor",
                display_name="Dashboard Mentor",
                email="dashboard.mentor@example.edu",
                phone="+919876543210",
            )
            mentor.set_password("MentorPass12")
            coordinator = User(
                username="dashboard_coordinator",
                role="coordinator",
                display_name="Dashboard Coordinator",
                email="dashboard.coordinator@example.edu",
                active=False,
            )
            coordinator.set_password("CoordPass12")
            team_leader = User(
                username="dashboard_leader",
                role="team_leader",
                display_name="Dashboard Team Leader",
                email="dashboard.leader@example.edu",
                phone="+919123456780",
                team_id=teams[0].id,
            )
            team_leader.set_password("LeaderPass12")
            judge = User(
                username="dashboard_judge",
                role="judge",
                display_name="Dashboard Judge",
                email="dashboard.judge@example.edu",
            )
            judge.set_password("JudgePass12")
            db.session.add_all([mentor, coordinator, team_leader, judge])
            db.session.flush()
            db.session.add_all(
                [
                    MentorTeamAssignment(mentor_id=mentor.id, team_id=teams[0].id),
                    CoordinatorTeamAssignment(
                        coordinator_id=coordinator.id,
                        team_id=teams[1].id,
                    ),
                ]
            )
            db.session.commit()

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        response = self.client.get("/admin")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Mentors", response.data)
        self.assertIn(b"Coordinators", response.data)
        self.assertLess(
            response.data.index(b"Mentors"),
            response.data.index(b"Team portfolio"),
        )
        self.assertIn(b"Dashboard Mentor", response.data)
        self.assertIn(b"Dashboard Coordinator", response.data)
        self.assertIn(b"dashboard_mentor", response.data)
        self.assertIn(b"dashboard_coordinator", response.data)
        self.assertIn(b"All user accounts", response.data)
        self.assertIn(b"Open dashboard as", response.data)
        self.assertIn(b"Open as user", response.data)
        self.assertIn(b"View-only", response.data)
        self.assertIn(b"Full access", response.data)
        self.assertIn(b"dashboard.leader@example.edu", response.data)
        self.assertIn(b"dashboard.judge@example.edu", response.data)
        self.assertIn(b"dashboard.mentor@example.edu", response.data)
        self.assertIn(b"+919123456780", response.data)
        self.assertIn(b"Inactive", response.data)
        self.assertIn(b"FT_01", response.data)
        self.assertIn(b"FT_02", response.data)

    def test_admin_can_download_complete_backup_with_database_exports_and_uploads(self):
        marker_path = os.path.join(self.upload_dir.name, "backup-export-marker.txt")
        with open(marker_path, "w", encoding="utf-8") as marker_file:
            marker_file.write("uploaded event file")
        try:
            anonymous_client = self.app.test_client()
            self.assertEqual(anonymous_client.get("/admin/data").status_code, 302)
            self.assertEqual(
                anonymous_client.get("/admin/data/backup").status_code,
                302,
            )
            self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
            data_page = self.client.get("/admin/data")
            self.assertEqual(data_page.status_code, 200)
            self.assertIn(b"Download complete backup", data_page.data)
            self.assertIn(b"Data & backups", self.client.get("/admin").data)
            response = self.client.get("/admin/data/backup")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.mimetype, "application/zip")
            with zipfile.ZipFile(io.BytesIO(response.data)) as backup:
                file_names = set(backup.namelist())
                self.assertIn(
                    "database/future_tech_innovision.sqlite",
                    file_names,
                )
                self.assertIn("database/schema.sql", file_names)
                self.assertIn("data/team.csv", file_names)
                self.assertIn("uploads/backup-export-marker.txt", file_names)
                manifest = json.loads(backup.read("manifest.json"))
                self.assertTrue(manifest["database_snapshot_included"])
                self.assertEqual(
                    manifest["password_hashes_in_user_csv"],
                    "redacted",
                )
                with tempfile.TemporaryDirectory() as snapshot_directory:
                    snapshot_path = os.path.join(
                        snapshot_directory, "database.sqlite"
                    )
                    with open(snapshot_path, "wb") as snapshot_file:
                        snapshot_file.write(
                            backup.read("database/future_tech_innovision.sqlite")
                        )
                    snapshot_connection = sqlite3.connect(snapshot_path)
                    try:
                        tables = {
                            row[0]
                            for row in snapshot_connection.execute(
                                "SELECT name FROM sqlite_master WHERE type='table'"
                            )
                        }
                        self.assertIn("team", tables)
                        self.assertIn("user", tables)
                    finally:
                        snapshot_connection.close()
                user_rows = list(
                    csv.DictReader(
                        io.StringIO(
                            backup.read("data/user.csv").decode("utf-8-sig")
                        )
                    )
                )
                admin_row = next(row for row in user_rows if row["username"] == "admin")
                self.assertEqual(admin_row["password_hash"], "")
        finally:
            if os.path.exists(marker_path):
                os.remove(marker_path)

    def test_domain_cards_open_rosters_and_admin_only_details(self):
        with self.app.app_context():
            domain = Domain.query.order_by(Domain.id).first()
            team = Team.query.filter_by(team_code="FT_01").one()
            team.domain_id = domain.id
            team.project_title = "Domain Details Project"
            team.leader_name = "Domain Leader"
            team.email = "leader-private@example.test"
            team.abstract = "Public project abstract"
            member = TeamMember(
                team_id=team.id,
                full_name="Domain Member",
                prn="PRIVATEPRN01",
                email="member-private@example.test",
            )
            mentor = User(
                username="domain_mentor",
                role="mentor",
                display_name="Assigned Domain Mentor",
            )
            mentor.set_password("MentorPassword123!")
            coordinator = User(
                username="domain_coordinator",
                role="coordinator",
                display_name="Assigned Domain Coordinator",
            )
            coordinator.set_password("CoordinatorPassword123!")
            judge = User(
                username="domain_judge",
                role="judge",
                display_name="Assigned Domain Judge",
            )
            judge.set_password("JudgePassword123!")
            db.session.add_all([member, mentor, coordinator, judge])
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id),
                CoordinatorDomainAssignment(
                    coordinator_id=coordinator.id,
                    domain_id=domain.id,
                ),
                JudgeAssignment(judge_id=judge.id, team_id=team.id),
                Evaluation(
                    team_id=team.id,
                    judge_id=judge.id,
                    total=33,
                    finalized=True,
                    approved=True,
                ),
                AdminEvaluation(team_id=team.id, total=36, finalized=True),
            ])
            domain_id = domain.id
            db.session.commit()

        home_page = self.client.get("/")
        self.assertIn(f"/domains/{domain_id}".encode(), home_page.data)
        public_page = self.client.get(f"/domains/{domain_id}")
        self.assertEqual(public_page.status_code, 200)
        self.assertIn(b"Domain Details Project", public_page.data)
        self.assertIn(b"Domain Leader", public_page.data)
        self.assertIn(b"Domain Member", public_page.data)
        self.assertNotIn(b"PRIVATEPRN01", public_page.data)
        self.assertNotIn(b"leader-private@example.test", public_page.data)
        self.assertNotIn(b"Assigned Domain Mentor", public_page.data)
        self.assertNotIn(b"33.0 / 60", public_page.data)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        admin_page = self.client.get(f"/domains/{domain_id}")
        self.assertIn(b"PRIVATEPRN01", admin_page.data)
        self.assertIn(b"leader-private@example.test", admin_page.data)
        self.assertIn(b"Assigned Domain Mentor", admin_page.data)
        self.assertIn(b"Assigned Domain Coordinator", admin_page.data)
        self.assertIn(b"Assigned Domain Judge", admin_page.data)
        self.assertIn(b"33.0 / 60", admin_page.data)

    def test_admin_creates_mentors_and_coordinators_with_one_or_two_selected_teams(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        accounts_page = self.client.get("/admin/users")
        self.assertEqual(accounts_page.status_code, 200)
        self.assertIn(b'id="staff-team-fields"', accounts_page.data)
        self.assertIn(b'name="team_ids"', accounts_page.data)
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02", "FT_03"])
            ).order_by(Team.team_code).all()
            team_ids = {team.team_code: team.id for team in teams}

        for role, username in (("mentor", "selected_mentor"), ("coordinator", "selected_coordinator")):
            response = self.client.post("/admin/users", data={
                "username": username,
                "display_name": username.replace("_", " ").title(),
                "role": role,
                "team_ids": [str(team_ids["FT_01"]), str(team_ids["FT_03"])],
                "password": "StaffPass12",
            })
            self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            mentor = User.query.filter_by(username="selected_mentor").one()
            coordinator = User.query.filter_by(username="selected_coordinator").one()
            self.assertEqual(
                {assignment.team_id for assignment in mentor.mentor_team_assignments},
                {team_ids["FT_01"], team_ids["FT_03"]},
            )
            self.assertEqual(
                {assignment.team_id for assignment in coordinator.coordinator_team_assignments},
                {team_ids["FT_01"], team_ids["FT_03"]},
            )

        single_team_mentor = self.client.post("/admin/users", data={
            "username": "single_team_mentor",
            "display_name": "Single Team Mentor",
            "role": "mentor",
            "team_ids": [str(team_ids["FT_02"])],
            "password": "MentorPasswordLongerThanTwelve123!",
        })
        self.assertEqual(single_team_mentor.status_code, 302)
        with self.app.app_context():
            mentor = User.query.filter_by(username="single_team_mentor").one()
            self.assertTrue(
                mentor.check_password("MentorPasswordLongerThanTwelve123!")
            )
            self.assertEqual(
                {assignment.team_id for assignment in mentor.mentor_team_assignments},
                {team_ids["FT_02"]},
            )

        assignments_page = self.client.get("/admin/role-assignments")
        self.assertIn(b"Assign one or two teams to each mentor", assignments_page.data)
        with self.app.app_context():
            mentor_id = User.query.filter_by(username="selected_mentor").one().id
        single_team_assignment = self.client.post(
            "/admin/role-assignments/mentor",
            data={
                "mentor_id": str(mentor_id),
                "team_ids": [str(team_ids["FT_02"])],
                "csrf_token": "test-token",
            },
        )
        self.assertEqual(single_team_assignment.status_code, 302)
        with self.app.app_context():
            mentor = db.session.get(User, mentor_id)
            self.assertEqual(
                {assignment.team_id for assignment in mentor.mentor_team_assignments},
                {team_ids["FT_02"]},
            )

        single_team_coordinator = self.client.post("/admin/users", data={
            "username": "single_team_coordinator",
            "display_name": "Single Team Coordinator",
            "role": "coordinator",
            "team_ids": [str(team_ids["FT_02"])],
            "password": "StaffPass12",
        })
        self.assertEqual(single_team_coordinator.status_code, 302)
        with self.app.app_context():
            coordinator = User.query.filter_by(
                username="single_team_coordinator"
            ).one()
            self.assertEqual(
                {assignment.team_id for assignment in coordinator.coordinator_team_assignments},
                {team_ids["FT_02"]},
            )

        missing_assignment = self.client.post("/admin/users", data={
            "username": "unassigned_mentor",
            "role": "mentor",
            "password": "StaffPass12",
        })
        self.assertEqual(missing_assignment.status_code, 302)
        with self.app.app_context():
            self.assertIsNone(User.query.filter_by(username="unassigned_mentor").first())

    def test_coordinator_can_be_assigned_one_or_two_teams_and_access_is_scoped(self):
        with self.app.app_context():
            teams = Team.query.filter(Team.team_code.in_(["FT_01", "FT_02", "FT_03"])).order_by(Team.team_code).all()
            for team, title in zip(teams, ("Assigned One", "Assigned Two", "Unassigned Three")):
                team.project_title = title
            coordinator = User(username="two_team_coordinator", role="coordinator")
            coordinator.set_password("CoordinatorPassword123!")
            db.session.add(coordinator)
            db.session.commit()
            coordinator_id = coordinator.id
            team_ids = [team.id for team in teams]

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        single = self.client.post(
            "/admin/role-assignments/coordinator",
            data={"coordinator_id": coordinator_id, "team_ids": [str(team_ids[0])]},
        )
        self.assertEqual(single.status_code, 302)
        with self.app.app_context():
            self.assertEqual(CoordinatorTeamAssignment.query.count(), 1)

        self.client.post("/logout")
        self.assertEqual(
            self.login("two_team_coordinator", "CoordinatorPassword123!").status_code,
            302,
        )
        dashboard = self.client.get("/role/Coordinator")
        self.assertIn(b"Assigned One", dashboard.data)
        self.assertNotIn(b"Assigned Two", dashboard.data)
        self.assertNotIn(b"Unassigned Three", dashboard.data)

        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)

        saved = self.client.post(
            "/admin/role-assignments/coordinator",
            data={"coordinator_id": coordinator_id, "team_ids": [str(team_ids[0]), str(team_ids[1])]},
        )
        self.assertEqual(saved.status_code, 302)
        self.client.post("/logout")
        self.assertEqual(
            self.login("two_team_coordinator", "CoordinatorPassword123!").status_code,
            302,
        )
        dashboard = self.client.get("/role/Coordinator")
        self.assertIn(b"Assigned One", dashboard.data)
        self.assertIn(b"Assigned Two", dashboard.data)
        self.assertNotIn(b"Unassigned Three", dashboard.data)

    def test_mentor_and_coordinator_dashboards_only_expose_assigned_records(self):
        with self.app.app_context():
            first_team = Team.query.filter_by(team_code="FT_01").one()
            second_team = Team.query.filter_by(team_code="FT_02").one()
            first_domain = Domain.query.order_by(Domain.id).first()
            first_team.domain_id = first_domain.id
            second_team.domain_id = first_domain.id
            first_team.project_title = "Assigned Project One"
            second_team.project_title = "Private Project Two"
            db.session.add_all(
                [
                    TeamMember(team_id=first_team.id, full_name="Assigned Student", prn="SCOPE001"),
                    TeamMember(team_id=second_team.id, full_name="Private Student", prn="SCOPE002"),
                ]
            )
            mentor = User(username="scoped_mentor", role="mentor", display_name="Assigned Mentor")
            mentor.set_password("MentorPassword123!")
            coordinator = User(username="scoped_coordinator", role="coordinator", display_name="Assigned Coordinator")
            coordinator.set_password("CoordinatorPassword123!")
            db.session.add_all([mentor, coordinator])
            db.session.flush()
            db.session.add(MentorTeamAssignment(mentor_id=mentor.id, team_id=first_team.id))
            db.session.add(
                CoordinatorTeamAssignment(
                    coordinator_id=coordinator.id,
                    team_id=first_team.id,
                )
            )
            db.session.commit()

        self.assertEqual(self.login("scoped_mentor", "MentorPassword123!").status_code, 302)
        mentor_view = self.client.get("/role/Mentor")
        self.assertIn(b"Assigned Project One", mentor_view.data)
        self.assertIn(b"Assigned Student", mentor_view.data)
        self.assertNotIn(b"Private Project Two", mentor_view.data)
        self.assertNotIn(b"Private Student", mentor_view.data)

        self.client.post("/logout")
        self.assertEqual(
            self.login("scoped_coordinator", "CoordinatorPassword123!").status_code,
            302,
        )
        coordinator_view = self.client.get("/role/Coordinator")
        self.assertIn(b"Assigned Project One", coordinator_view.data)
        self.assertNotIn(b"Private Project Two", coordinator_view.data)

    def test_project_videos_are_only_visible_to_admin_and_assigned_roles(self):
        with self.app.app_context():
            first_team = Team.query.filter_by(team_code="FT_01").one()
            second_team = Team.query.filter_by(team_code="FT_02").one()
            first_domain = Domain.query.order_by(Domain.id).first()
            first_team.domain_id = first_domain.id
            second_team.domain_id = first_domain.id
            mentor = User(username="video_mentor", role="mentor", display_name="Mentor")
            mentor.set_password("MentorPassword123!")
            coordinator = User(username="video_coordinator", role="coordinator", display_name="Coordinator")
            coordinator.set_password("CoordinatorPassword123!")
            db.session.add_all([mentor, coordinator])
            db.session.flush()
            db.session.add(MentorTeamAssignment(mentor_id=mentor.id, team_id=first_team.id))
            db.session.add(
                CoordinatorTeamAssignment(
                    coordinator_id=coordinator.id,
                    team_id=first_team.id,
                )
            )
            db.session.add_all(
                [
                    ProjectVideo(team_id=first_team.id, filename="first.mp4"),
                    ProjectVideo(team_id=second_team.id, filename="second.mp4"),
                ]
            )
            db.session.commit()
            first_id = ProjectVideo.query.filter_by(team_id=first_team.id).one().id
            second_id = ProjectVideo.query.filter_by(team_id=second_team.id).one().id
            video_dir = os.path.join(self.upload_dir.name, "videos")
            os.makedirs(video_dir, exist_ok=True)
            with open(os.path.join(video_dir, "first.mp4"), "wb") as file:
                file.write(b"assigned video")
            with open(os.path.join(video_dir, "second.mp4"), "wb") as file:
                file.write(b"private video")

        self.assertEqual(self.login("video_mentor", "MentorPassword123!").status_code, 302)
        response = self.client.get(f"/team/video/{first_id}")
        self.assertEqual(response.status_code, 200)
        response.close()
        self.assertEqual(self.client.get(f"/team/video/{second_id}").status_code, 403)
        self.client.post("/logout")
        self.assertEqual(
            self.login("video_coordinator", "CoordinatorPassword123!").status_code,
            302,
        )
        response = self.client.get(f"/team/video/{first_id}")
        self.assertEqual(response.status_code, 200)
        response.close()
        self.assertEqual(self.client.get(f"/team/video/{second_id}").status_code, 403)
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        response = self.client.get(f"/team/video/{second_id}")
        self.assertEqual(response.status_code, 200)
        response.close()

    def test_team_leader_can_upload_video_during_registration_and_assigned_users_can_watch(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            member_user = User(
                username="early_video_member",
                role="team_member",
                display_name="Team Member",
                team_id=team.id,
            )
            member_user.set_password("TeamMemberPassword123!")
            member = TeamMember(
                team_id=team.id,
                full_name="Team Member",
                prn="EARLYVIDEO001",
                user=member_user,
            )
            mentor = User(username="early_video_mentor", role="mentor")
            mentor.set_password("MentorPassword123!")
            coordinator = User(username="early_video_coordinator", role="coordinator")
            coordinator.set_password("CoordinatorPassword123!")
            db.session.add_all([member, mentor, coordinator])
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id),
                CoordinatorTeamAssignment(coordinator_id=coordinator.id, team_id=team.id),
            ])
            db.session.commit()

        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        team_page = self.client.get("/team")
        self.assertEqual(team_page.status_code, 200)
        self.assertIn(b"Upload or replace your MP4 presentation", team_page.data)
        self.assertIn(b"Team name", team_page.data)
        self.assertNotIn(b"Project presentation video (optional)", team_page.data)
        response = self.client.post("/team/update", data={
            "name": "Registration Team Name",
            "project_title": "Project with video",
        })
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            self.assertEqual(team.name, "Registration Team Name")
            self.assertEqual(team.email, "")
            self.assertEqual(ProjectVideo.query.filter_by(team_id=team.id).count(), 0)
        response = self.client.post("/team/video", data={
            "video": (io.BytesIO(b"\x00\x00\x00\x18ftypisomregistration video"), "demo.mp4"),
        }, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            video = ProjectVideo.query.join(Team).filter(Team.team_code == "FT_01").one()
            video_id = video.id
            previous_filename = video.filename
            self.assertEqual(video.team.project_title, "Project with video")
        replacement = self.client.post("/team/video", data={
            "video": (io.BytesIO(b"\x00\x00\x00\x18ftypisomreplacement video"), "replacement.mp4"),
        }, content_type="multipart/form-data")
        self.assertEqual(replacement.status_code, 302)
        with self.app.app_context():
            video = ProjectVideo.query.join(Team).filter(Team.team_code == "FT_01").one()
            self.assertEqual(video.id, video_id)
            self.assertNotEqual(video.filename, previous_filename)
            self.assertFalse(os.path.exists(os.path.join(
                self.upload_dir.name, "videos", previous_filename
            )))

        for username, password, role, dashboard in (
            ("early_video_member", "TeamMemberPassword123!", "team_member", "/member"),
            ("early_video_mentor", "MentorPassword123!", "mentor", "/role/Mentor"),
            ("early_video_coordinator", "CoordinatorPassword123!", "coordinator", "/role/Coordinator"),
        ):
            viewer = self.app.test_client()
            self.assertEqual(viewer.post("/login", data={
                "username": username,
                "password": password,
                "role": role,
            }).status_code, 302)
            page = viewer.get(dashboard)
            self.assertEqual(page.status_code, 200)
            self.assertIn(f"/team/video/{video_id}".encode(), page.data)
            playback = viewer.get(f"/team/video/{video_id}")
            self.assertEqual(playback.status_code, 200)
            playback.close()

    def test_team_submissions_support_ppt_and_documentation_with_team_scoped_downloads(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            other_team = Team.query.filter_by(team_code="FT_02").one()
            member_user = User(
                username="submission_member",
                role="team_member",
                display_name="Team Member",
                team_id=team.id,
            )
            member_user.set_password("TeamMemberPassword123!")
            member = TeamMember(
                team_id=team.id,
                full_name="Team Member",
                prn="SUBMISSION001",
                user=member_user,
            )
            mentor = User(username="submission_mentor", role="mentor")
            mentor.set_password("MentorPassword123!")
            coordinator = User(username="submission_coordinator", role="coordinator")
            coordinator.set_password("CoordinatorPassword123!")
            other_leader = User.query.filter_by(team_id=other_team.id).first()
            db.session.add_all([member, mentor, coordinator])
            db.session.flush()
            db.session.add_all([
                MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id),
                CoordinatorTeamAssignment(
                    coordinator_id=coordinator.id,
                    team_id=team.id,
                ),
            ])
            db.session.commit()

        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        for submission_type, filename, content in (
            ("ppt", "presentation.pptx", b"PPTX content"),
            ("documentation", "project-report.pdf", b"PDF content"),
        ):
            response = self.client.post(
                f"/team/submissions/{submission_type}",
                data={"file": (io.BytesIO(content), filename)},
                content_type="multipart/form-data",
            )
            self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            submissions = TeamSubmission.query.join(Team).filter(
                Team.team_code == "FT_01"
            ).order_by(TeamSubmission.submission_type).all()
            self.assertEqual(
                {submission.submission_type for submission in submissions},
                {"ppt", "documentation"},
            )
            submission_ids = {
                submission.submission_type: submission.id for submission in submissions
            }
            submission_filenames = {
                submission.submission_type: submission.filename
                for submission in submissions
            }

        team_page = self.client.get("/team")
        self.assertIn(b"PPT submission", team_page.data)
        self.assertIn(b"Project documentation", team_page.data)
        self.assertIn(b"presentation.pptx", team_page.data)
        self.assertIn(b"project-report.pdf", team_page.data)

        for username, password, role, dashboard in (
            ("submission_member", "TeamMemberPassword123!", "team_member", "/member"),
            ("submission_mentor", "MentorPassword123!", "mentor", "/role/Mentor"),
            ("submission_coordinator", "CoordinatorPassword123!", "coordinator", "/role/Coordinator"),
        ):
            viewer = self.app.test_client()
            self.assertEqual(viewer.post("/login", data={
                "username": username,
                "password": password,
                "role": role,
            }).status_code, 302)
            page = viewer.get(dashboard)
            self.assertEqual(page.status_code, 200)
            for submission_type in ("ppt", "documentation"):
                download = viewer.get(
                    f"/team/submissions/{submission_ids[submission_type]}"
                )
                self.assertEqual(download.status_code, 200)
                self.assertEqual(
                    download.data,
                    b"PPTX content" if submission_type == "ppt" else b"PDF content",
                )
                download.close()

        other_viewer = self.app.test_client()
        self.assertEqual(other_viewer.post("/login", data={
            "username": "FT_02",
            "password": "ChangeMe123!",
            "role": "team_leader",
        }).status_code, 302)
        self.assertEqual(
            other_viewer.get(
                f"/team/submissions/{submission_ids['ppt']}"
            ).status_code,
            403,
        )
        self.assertEqual(
            self.client.post(
                "/team/submissions/ppt",
                data={"file": (io.BytesIO(b"bad type"), "report.pdf")},
                content_type="multipart/form-data",
            ).status_code,
            302,
        )
        with self.app.app_context():
            submission = db.session.get(
                TeamSubmission, submission_ids["ppt"]
            )
            self.assertEqual(submission.filename, submission_filenames["ppt"])

        admin_client = self.app.test_client()
        self.assertEqual(admin_client.post("/login", data={
            "username": "admin",
            "password": "Admin@12345!",
            "role": "admin",
        }).status_code, 302)
        admin_monitor = admin_client.get("/admin/monitor")
        self.assertEqual(admin_monitor.status_code, 200)
        self.assertIn(b"Download PPT", admin_monitor.data)
        self.assertIn(b"Download project documentation", admin_monitor.data)
        for submission_type in ("ppt", "documentation"):
            admin_download = admin_client.get(
                f"/team/submissions/{submission_ids[submission_type]}"
            )
            self.assertEqual(admin_download.status_code, 200)
            admin_download.close()
        task_page = admin_client.get("/tasks")
        self.assertEqual(task_page.status_code, 200)
        self.assertIn(b"Email: Not configured", task_page.data)
        self.assertIn(b"SMS: Not configured", task_page.data)
        self.assertIn(b"Manage all accounts", admin_monitor.data)

    def test_seeded_team_accounts_receive_separate_random_passwords_by_default(self):
        configured_password = self.app.config["TEAM_INITIAL_PASSWORD"]
        self.app.config["TEAM_INITIAL_PASSWORD"] = None
        generated = [f"RandFT{index:02d}Pass" for index in range(27)]
        try:
            with self.app.app_context():
                db.drop_all()
                db.create_all()
                with patch("app.secrets.token_urlsafe", side_effect=generated):
                    seed_initial_data()
                first = User.query.filter_by(username="FT_01").one()
                second = User.query.filter_by(username="FT_02").one()
                self.assertTrue(first.check_password(generated[0]))
                self.assertTrue(second.check_password(generated[1]))
                self.assertFalse(second.check_password(generated[0]))
        finally:
            self.app.config["TEAM_INITIAL_PASSWORD"] = configured_password

    def test_password_change_requires_four_characters_without_a_maximum(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        long_password = "PasswordLongerThanTwelveCharacters123!"
        response = self.client.post(
            "/account/password",
            data={
                "current_password": "Admin@12345!",
                "new_password": long_password,
                "confirm_password": long_password,
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            self.assertTrue(User.query.filter_by(username="admin").one().check_password(long_password))
        too_short = self.client.post(
            "/account/password",
            data={
                "current_password": long_password,
                "new_password": "abc",
                "confirm_password": "abc",
            },
        )
        self.assertEqual(too_short.status_code, 200)
        too_long = self.client.post(
            "/account/password",
            data={
                "current_password": long_password,
                "new_password": "PasswordIsTooLong",
                "confirm_password": "PasswordIsTooLong",
            },
        )
        self.assertEqual(too_long.status_code, 302)
        with self.app.app_context():
            self.assertTrue(
                User.query.filter_by(username="admin").one().check_password(
                    "PasswordIsTooLong"
                )
            )
        self.assertEqual(self.client.get("/admin").status_code, 200)

    def test_admin_password_reset_invalidates_the_users_existing_session(self):
        team_client = self.app.test_client()
        self.assertEqual(
            team_client.post(
                "/login",
                data={"username": "FT_01", "password": "ChangeMe123!", "role": "team_leader"},
            ).status_code,
            302,
        )
        self.assertEqual(team_client.get("/team").status_code, 200)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            team_user = User.query.filter_by(username="FT_01").one()
            team_user_id = team_user.id
        reset_response = self.client.post(
            f"/admin/users/{team_user_id}/password",
            data={"password": "FreshPasswordLongerThanTwelve123!"},
        )
        self.assertEqual(reset_response.status_code, 302)
        self.assertEqual(team_client.get("/team").status_code, 302)
        refreshed_client = self.app.test_client()
        self.assertEqual(
            refreshed_client.post(
                "/login",
                data={
                    "username": "FT_01",
                    "password": "FreshPasswordLongerThanTwelve123!",
                    "role": "team_leader",
                },
            ).status_code,
            302,
        )

    def test_login_throttles_repeated_failed_attempts(self):
        for _ in range(10):
            self.assertEqual(self.login("missing_user", "wrong").status_code, 200)
        self.assertEqual(self.login("missing_user", "wrong").status_code, 429)

    def test_login_requires_the_selected_role_to_match_account(self):
        response = self.client.get("/login")
        self.assertIn(b'name="role"', response.data)
        self.assertIn(b'data-password-toggle', response.data)
        self.assertIn(b'data-password-toggle-button', response.data)
        self.assertEqual(
            self.client.post(
                "/login",
                data={
                    "username": "admin",
                    "password": "Admin@12345!",
                    "role": "mentor",
                },
            ).status_code,
            200,
        )
        self.assertEqual(
            self.client.post(
                "/login",
                data={
                    "username": "admin",
                    "password": "Admin@12345!",
                    "role": "admin",
                },
            ).status_code,
            302,
        )

    def test_public_login_and_authenticated_pages_keep_the_same_background_theme(self):
        self.assertIn(b'<body class="dashboard-space">', self.client.get("/").data)
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.assertIn(
            b'<body class="dashboard-space">',
            self.client.get("/admin").data,
        )
        self.assertEqual(self.client.post("/logout").status_code, 302)
        self.assertIn(
            b'<body class="dashboard-space">',
            self.client.get("/login").data,
        )

    def test_admin_team_management_displays_only_final_approved_scores(self):
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02"])
            ).order_by(Team.team_code).all()
            db.session.add_all(
                [
                    AdminEvaluation(
                        team_id=teams[0].id,
                        total=34,
                        finalized=True,
                    ),
                    AdminEvaluation(
                        team_id=teams[1].id,
                        total=25,
                        finalized=False,
                    ),
                ]
            )
            db.session.commit()
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        response = self.client.get("/admin/teams")
        self.assertIn(b"Approved 40 marks", response.data)
        self.assertIn(b"34.0 / 40", response.data)
        self.assertIn(b"Awaiting approval", response.data)

    def test_rankings_are_not_public_before_publication(self):
        response = self.client.get("/rankings")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"The final leaderboard will be announced here by the admin", response.data)
        self.assertNotIn(b"Total Marks / 100", response.data)
        visitor = self.app.test_client()
        self.assertEqual(visitor.post("/admin/publish-leaderboard").status_code, 302)

    def test_nominations_include_all_eligible_teams_when_domain_has_fewer_than_three(self):
        eligible_counts = [2, 3, 3, 3, 3]
        team_counts = [3, 3, 3, 3, 3]
        with self.app.app_context():
            judge = User(username="ranking_judge", role="judge", display_name="Judge")
            judge.set_password("RankingJudgePassword123!")
            db.session.add(judge)
            db.session.flush()
            team_number = 1
            for domain_index, (domain, team_count) in enumerate(
                zip(Domain.query.order_by(Domain.id), team_counts)
            ):
                for position in range(team_count):
                    team = Team.query.filter_by(team_code=f"FT_{team_number:02d}").one()
                    team.domain_id = domain.id
                    team.registration_status = "Approved"
                    if domain_index == 0 and position == 2:
                        team_number += 1
                        continue
                    db.session.add(JudgeAssignment(judge_id=judge.id, team_id=team.id))
                    db.session.add(
                        Evaluation(
                            team_id=team.id,
                            judge_id=judge.id,
                            total=60 - position,
                            finalized=True,
                            approved=True,
                        )
                    )
                    db.session.add(
                        AdminEvaluation(
                            team_id=team.id,
                            development=10,
                            teamwork=10,
                            documentation=8,
                            discipline=6,
                            feasibility=6,
                            total=40,
                            finalized=True,
                        )
                    )
                    team_number += 1
            db.session.commit()

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.assertEqual(self.client.post("/admin/nominate").status_code, 302)
        with self.app.app_context():
            nominations = Nomination.query.all()
            self.assertEqual(len(nominations), sum(eligible_counts))
            for domain, allocation in zip(
                Domain.query.order_by(Domain.id), eligible_counts
            ):
                domain_nominations = [
                    nomination for nomination in nominations
                    if nomination.team.domain_id == domain.id
                ]
                self.assertEqual(len(domain_nominations), allocation)
                self.assertEqual(
                    [nomination.rank for nomination in sorted(domain_nominations, key=lambda row: row.rank)],
                    list(range(1, allocation + 1)),
                )

        self.assertEqual(self.client.post("/admin/publish-nominations").status_code, 302)
        rankings = self.client.get("/rankings")
        self.assertEqual(rankings.status_code, 200)
        self.assertIn(b"Admin preview only", rankings.data)
        self.assertIn(b"Total Marks / 100", rankings.data)
        self.assertNotIn(b"Final Winners &amp; Leaderboard Announced", rankings.data)

        with self.app.app_context():
            domain_winner_ids = [
                Nomination.query.join(Team).filter(
                    Team.domain_id == domain.id,
                    Nomination.rank == 1,
                    Nomination.published.is_(True),
                ).one().team_id
                for domain in Domain.query.order_by(Domain.id)
            ]
        for team_id in domain_winner_ids:
            self.assertEqual(self.client.post(
                "/admin/winner",
                data={"team_id": str(team_id)},
            ).status_code, 302)

        self.client.post("/logout")
        public_before_announcement = self.app.test_client()
        rankings = public_before_announcement.get("/rankings")
        self.assertIn(b"The final leaderboard will be announced here by the admin", rankings.data)
        self.assertNotIn(b"Total Marks / 100", rankings.data)
        self.assertEqual(public_before_announcement.get("/announcements").status_code, 200)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.assertEqual(self.client.post("/admin/publish-leaderboard").status_code, 302)
        rankings = self.client.get("/rankings")
        self.assertIn(b"The event admin has announced the final winners", rankings.data)
        self.assertIn(b"Winner", rankings.data)
        announcement_page = self.client.get("/announcements")
        self.assertIn(b"Final Winners &amp; Leaderboard Announced", announcement_page.data)
        self.assertEqual(self.client.post("/admin/approve/1", data={"status": "Rejected"}).status_code, 409)

    def test_nominations_include_eligible_teams_when_other_domains_are_empty(self):
        with self.app.app_context():
            domain = Domain.query.order_by(Domain.id).first()
            judge = User(username="ineligible_judge", role="judge", display_name="Judge")
            judge.set_password("IneligibleJudgePassword123!")
            db.session.add(judge)
            db.session.flush()
            for team_code in ("FT_01", "FT_02"):
                team = Team.query.filter_by(team_code=team_code).one()
                team.domain_id = domain.id
                team.registration_status = "Approved"
                db.session.add(JudgeAssignment(judge_id=judge.id, team_id=team.id))
                db.session.add(
                    Evaluation(
                        team_id=team.id,
                        judge_id=judge.id,
                        total=60,
                        finalized=True,
                        approved=True,
                    )
                )
                db.session.add(
                    AdminEvaluation(
                        team_id=team.id,
                        development=10,
                        teamwork=10,
                        documentation=8,
                        discipline=6,
                        feasibility=6,
                        total=40,
                        finalized=True,
                    )
                )
            db.session.commit()
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.assertEqual(self.client.post("/admin/nominate").status_code, 302)
        with self.app.app_context():
            nominations = Nomination.query.order_by(Nomination.rank).all()
            self.assertEqual(len(nominations), 2)
            self.assertEqual([item.rank for item in nominations], [1, 2])
            self.assertEqual(
                {item.team.team_code for item in nominations},
                {"FT_01", "FT_02"},
            )

    def test_mentor_score_requires_admin_approval_and_staff_only_visibility(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            mentor = User(
                username="score_mentor",
                role="mentor",
                display_name="Score Mentor",
            )
            mentor.set_password("Mentor123")
            coordinator = User(
                username="score_coordinator",
                role="coordinator",
                display_name="Score Coordinator",
            )
            coordinator.set_password("Coord123")
            member = User(
                username="score_member",
                role="team_member",
                display_name="Score Member",
                team_id=team.id,
            )
            member.set_password("Member123")
            db.session.add_all([mentor, coordinator, member])
            db.session.flush()
            db.session.add(MentorTeamAssignment(mentor_id=mentor.id, team_id=team.id))
            db.session.add(CoordinatorTeamAssignment(
                coordinator_id=coordinator.id,
                team_id=team.id,
            ))
            db.session.commit()
            team_id = team.id

        self.assertEqual(self.login("score_mentor", "Mentor123").status_code, 302)
        proposal_response = self.client.post(
            f"/mentor/evaluate/{team_id}",
            data={
                "development": "9",
                "teamwork": "8",
                "documentation": "7",
                "discipline": "6",
                "feasibility": "5",
            },
        )
        self.assertEqual(proposal_response.status_code, 302)
        with self.app.app_context():
            proposal = MentorEvaluation.query.one()
            self.assertEqual(proposal.status, "Pending")
            self.assertEqual(proposal.total, 35)
            self.assertIsNone(AdminEvaluation.query.filter_by(team_id=team_id).first())

        mentor_dashboard = self.client.get("/role/Mentor")
        self.assertNotIn(b"35.0 / 40", mentor_dashboard.data)
        self.assertNotIn(b"Approved evaluation marks", mentor_dashboard.data)
        self.client.post("/logout")
        self.assertEqual(self.login("score_coordinator", "Coord123").status_code, 302)
        coordinator_dashboard = self.client.get("/role/Coordinator")
        self.assertNotIn(b"35.0 / 40", coordinator_dashboard.data)
        self.assertNotIn(b"Mentor proposal", coordinator_dashboard.data)

        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            proposal_id = MentorEvaluation.query.one().id
        review_response = self.client.post(
            f"/admin/admin-evaluation/{team_id}",
            data={
                "mentor_evaluation_id": str(proposal_id),
                "development": "8",
                "teamwork": "8",
                "documentation": "7",
                "discipline": "5",
                "feasibility": "5",
            },
        )
        self.assertEqual(review_response.status_code, 302)
        with self.app.app_context():
            final_score = AdminEvaluation.query.filter_by(team_id=team_id).one()
            proposal = MentorEvaluation.query.one()
            self.assertTrue(final_score.finalized)
            self.assertEqual(final_score.total, 33)
            self.assertEqual(proposal.status, "Approved")
            self.assertEqual(proposal.reviewed_by, User.query.filter_by(username="admin").one().id)

        self.client.post("/logout")
        self.assertEqual(self.login("score_coordinator", "Coord123").status_code, 302)
        coordinator_dashboard = self.client.get("/role/Coordinator")
        self.assertIn(b"33.0 / 40", coordinator_dashboard.data)
        self.client.post("/logout")
        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        team_dashboard = self.client.get("/team")
        self.assertNotIn(b"33.0 / 40", team_dashboard.data)
        self.client.post("/logout")
        self.assertEqual(self.login("score_member", "Member123").status_code, 302)
        member_dashboard = self.client.get("/member")
        self.assertNotIn(b"33.0 / 40", member_dashboard.data)

    def test_admin_can_create_admin_and_add_or_safely_delete_teams(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        account_response = self.client.post(
            "/admin/users",
            data={
                "username": "second_admin",
                "display_name": "Second Administrator",
                "role": "admin",
                "password": "Admin1234",
            },
        )
        self.assertEqual(account_response.status_code, 302)
        with self.app.app_context():
            self.assertEqual(User.query.filter_by(username="second_admin", role="admin").count(), 1)

        create_response = self.client.post(
            "/admin/teams/create",
            data={
                "team_code": "FT_28",
                "name": "New Team",
                "leader_name": "New Leader",
            },
        )
        self.assertEqual(create_response.status_code, 302)
        with self.app.app_context():
            new_team = Team.query.filter_by(team_code="FT_28").one()
            leader = User.query.filter_by(username="FT_28").one()
            self.assertEqual(leader.role, "team_leader")
            self.assertEqual(leader.team_id, new_team.id)
            new_team_id = new_team.id
        self.assertEqual(
            self.client.post(f"/admin/teams/{new_team_id}/delete").status_code,
            302,
        )
        with self.app.app_context():
            self.assertIsNone(Team.query.filter_by(team_code="FT_28").first())
            self.assertIsNone(User.query.filter_by(username="FT_28").first())
            protected_team = Team(team_code="FT_29", name="Scored Team")
            db.session.add(protected_team)
            db.session.flush()
            db.session.add(AdminEvaluation(team_id=protected_team.id, finalized=True))
            db.session.commit()
            protected_team_id = protected_team.id
        self.client.post(f"/admin/teams/{protected_team_id}/delete")
        with self.app.app_context():
            self.assertIsNotNone(db.session.get(Team, protected_team_id))

    def test_domain_results_publish_individually_and_notify_team_stakeholders(self):
        with self.app.app_context():
            domain = Domain.query.order_by(Domain.id).first()
            judge = User(username="domain_judge", role="judge")
            judge.set_password("Judge123")
            mentor = User(username="domain_mentor", role="mentor")
            mentor.set_password("Mentor123")
            coordinator = User(username="domain_coordinator", role="coordinator")
            coordinator.set_password("Coord123")
            member = User(
                username="domain_member",
                role="team_member",
                display_name="Domain Member",
                team_id=Team.query.filter_by(team_code="FT_01").one().id,
            )
            member.set_password("Member123")
            db.session.add_all([judge, mentor, coordinator, member])
            db.session.flush()
            for team_code in ("FT_01", "FT_02"):
                team = Team.query.filter_by(team_code=team_code).one()
                team.domain_id = domain.id
                team.registration_status = "Approved"
                db.session.add(JudgeAssignment(judge_id=judge.id, team_id=team.id))
                db.session.add(Evaluation(
                    team_id=team.id,
                    judge_id=judge.id,
                    total=60,
                    finalized=True,
                    approved=True,
                ))
                db.session.add(AdminEvaluation(
                    team_id=team.id,
                    total=40,
                    finalized=True,
                ))
            ineligible_team = Team.query.filter_by(team_code="FT_03").one()
            ineligible_team.domain_id = domain.id
            ineligible_team.registration_status = "Approved"
            first_team = Team.query.filter_by(team_code="FT_01").one()
            db.session.add(MentorTeamAssignment(
                mentor_id=mentor.id,
                team_id=first_team.id,
            ))
            db.session.add(CoordinatorTeamAssignment(
                coordinator_id=coordinator.id,
                team_id=first_team.id,
            ))
            db.session.commit()
            domain_id = domain.id

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.client.post("/admin/nominate")
        self.client.post("/admin/publish-nominations")
        self.client.post("/logout")
        public_client = self.app.test_client()
        hidden_results = public_client.get("/rankings")
        self.assertNotIn(b"FT_01", hidden_results.data)
        self.assertEqual(
            self.login("admin", "Admin@12345!").status_code,
            302,
        )
        self.client.post(f"/admin/publish-domain-result/{domain_id}")
        with self.app.app_context():
            self.assertEqual(DomainResultPublication.query.filter_by(domain_id=domain_id).count(), 1)
            self.assertGreaterEqual(
                UserNotification.query.filter_by(title=f"{domain.code} results announced").count(),
                3,
            )
            unselected_leader = User.query.filter_by(username="FT_03").one()
            nomination_notice = UserNotification.query.filter_by(
                user_id=unselected_leader.id,
                title="Domain nominations published",
            ).one()
            self.assertIn("was not selected", nomination_notice.body)
        self.client.post("/logout")
        visible_results = public_client.get("/rankings")
        self.assertIn(b"FT_01", visible_results.data)
        self.assertNotIn(b"FT_03", visible_results.data)
        self.assertEqual(self.login("FT_01", "ChangeMe123!").status_code, 302)
        self.assertIn(b"Published ranking", self.client.get("/team").data)
        self.client.post("/logout")
        self.assertEqual(self.login("domain_member", "Member123").status_code, 302)
        self.assertIn(b"Published ranking", self.client.get("/member").data)
        self.client.post("/logout")
        self.assertEqual(self.login("domain_mentor", "Mentor123").status_code, 302)
        self.assertIn(b"Published ranking", self.client.get("/role/Mentor").data)
        self.client.post("/logout")
        self.assertEqual(self.login("domain_coordinator", "Coord123").status_code, 302)
        self.assertIn(b"Published ranking", self.client.get("/role/Coordinator").data)

    def test_certificate_verification_url_uses_canonical_public_domain(self):
        with patch.dict(
            self.app.config,
            {"CERTIFICATE_PUBLIC_BASE_URL": "https://certificates.example.org/"},
        ):
            with self.app.test_request_context("/"):
                self.assertEqual(
                    certificate_verification_url("FTI26-EXAMPLE"),
                    "https://certificates.example.org/certificate/FTI26-EXAMPLE",
                )

    def test_certificate_data_persists_when_initial_data_is_seeded(self):
        with self.app.app_context():
            certificate_type = CertificateType.query.filter_by(
                name="Certificate of Participation"
            ).one()
            template = CertificateTemplate(
                certificate_type=certificate_type,
                event_name="FUTURE-TECH INNOVISION 2026",
                event_date="7 November 2026",
                recipient_role="team_member",
                background_file="temporary.png",
            )
            certificate = Certificate(
                certificate_id="PERSIST-CERTIFICATE",
                type=certificate_type,
                template=template,
                recipient_name="Persistent Recipient",
            )
            db.session.add(certificate)
            db.session.commit()
            seed_initial_data()
            self.assertEqual(Certificate.query.count(), 1)
            self.assertEqual(CertificateTemplate.query.count(), 1)
            self.assertEqual(CertificateType.query.count(), 5)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        health = self.client.get("/healthz")
        self.assertEqual(health.status_code, 200)
        self.assertEqual(health.json, {"status": "ok"})
        certificate_page = self.client.get("/admin/certificates")
        self.assertEqual(certificate_page.status_code, 200)
        for type_name in (
            "Certificate of Participation",
            "Certificate of Winner",
            "Certificate of Appreciation - Coordinator",
            "Certificate of Appreciation - Mentor",
            "Certificate of Appreciation - Club Head and Judges",
        ):
            self.assertIn(type_name.encode(), certificate_page.data)
        self.assertEqual(
            self.client.get("/admin/certificate-template-starter").status_code,
            200,
        )

    def test_demo_certificate_is_dispatched_listed_and_deletable(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            domain = Domain.query.order_by(Domain.id).first()
            team.domain_id = domain.id
            team.project_title = "Demo smart project"
            student = User(
                username="demo_student",
                role="team_member",
                display_name="Demo Student",
                team_id=team.id,
                email="demo_student@example.edu",
            )
            student.set_password("Student123")
            db.session.add(student)
            db.session.flush()
            db.session.add(TeamMember(
                team_id=team.id,
                full_name="Demo Student",
                prn="DEMO-STUDENT-PRN",
                user_id=student.id,
                email=student.email,
            ))
            db.session.commit()
            student_id = student.id
            expected_domain_name = domain.name

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        image_buffer = io.BytesIO()
        Image.new("RGB", (842, 595), "white").save(image_buffer, format="PNG")
        image_buffer.seek(0)
        template_response = self.client.post(
            "/admin/certificates",
            data={
                "action": "template",
                "event_name": "FUTURE-TECH INNOVISION 2026",
                "event_date": "7 November 2026",
                "event_venue": "N K Orchid College",
                "type_name": "Certificate of Participation",
                "recipient_role": "team_member",
                "positions_json": "{}",
                "template_file": (image_buffer, "demo-background.png"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(
            template_response.status_code,
            302,
            template_response.get_data(as_text=True),
        )
        with self.app.app_context():
            template = CertificateTemplate.query.one()
            template_id = template.id
        self.assertEqual(
            self.client.post(
                f"/admin/certificate-templates/{template_id}/approve"
            ).status_code,
            302,
        )
        page = self.client.get("/admin/certificates")
        self.assertIn(b"Generate and send a demo certificate", page.data)
        self.assertIn(b"demo_student", page.data)
        demo_response = self.client.post(
            "/admin/certificates",
            data={
                "action": "demo_issue",
                "template_id": str(template_id),
                "user_id": str(student_id),
            },
        )
        self.assertEqual(demo_response.status_code, 302)
        with self.app.app_context():
            certificate = Certificate.query.one()
            certificate_db_id = certificate.id
            certificate_id = certificate.certificate_id
            self.assertEqual(certificate.user_id, student_id)
            self.assertEqual(certificate.team_code, "FT_01")
            self.assertEqual(certificate.dispatch_status, "dispatched")
            fields = json.loads(certificate.fields_json)
            self.assertEqual(fields["domain_name"], expected_domain_name)
            self.assertTrue(UserNotification.query.filter_by(
                user_id=student_id,
                title="Demo certificate dispatched: Certificate of Participation",
            ).first())
        issued_page = self.client.get("/admin/certificates")
        self.assertIn(certificate_id.encode(), issued_page.data)
        self.assertIn(b"Delete this issued certificate?", issued_page.data)
        self.client.post("/logout")
        self.assertEqual(self.login("demo_student", "Student123").status_code, 302)
        self.assertIn(
            certificate_id.encode(),
            self.client.get("/member").data,
        )
        pdf_response = self.client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(pdf_response.status_code, 200)
        self.assertIn(
            certificate_id,
            pymupdf.open(stream=pdf_response.data, filetype="pdf")[0].get_text(),
        )
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        deletion = self.client.post(
            f"/admin/certificates/{certificate_db_id}/delete"
        )
        self.assertEqual(deletion.status_code, 302)
        with self.app.app_context():
            self.assertIsNone(db.session.get(Certificate, certificate_db_id))
        self.assertNotIn(
            f"/certificate/{certificate_id}/download".encode(),
            self.client.get("/admin/certificates").data,
        )

    def test_certificate_template_csv_mapping_download_and_admin_delete(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            member = User(
                username="csv_student",
                role="team_member",
                display_name="CSV Student",
                team_id=team.id,
            )
            member.set_password("Student123")
            db.session.add(member)
            db.session.commit()
            team_id = team.id
            member_user_id = member.id

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        starter = self.client.get(
            "/admin/certificate-template-starter"
            "?type_name=Certificate%20of%20Participation"
        )
        self.assertEqual(starter.status_code, 200)
        starter_document = zipfile.ZipFile(io.BytesIO(starter.data))
        self.assertIn(
            b"&lt;name&gt;",
            starter_document.read("word/document.xml"),
        )
        certificate_page = self.client.get("/admin/certificates")
        self.assertEqual(certificate_page.status_code, 200)
        self.assertIn(b'name="template_docx"', certificate_page.data)
        self.assertIn(b"Word source file (optional)", certificate_page.data)
        self.assertIn(b"data-background-picker", certificate_page.data)
        self.assertIn(b"data-separate-background-toggle", certificate_page.data)
        self.assertIn(b"data-use-separate-background", certificate_page.data)
        self.assertIn(b"data-add-text", certificate_page.data)
        self.assertIn(b"data-text-color", certificate_page.data)
        self.assertIn(b"data-field-style-target", certificate_page.data)
        self.assertIn(b"data-field-style-color", certificate_page.data)
        self.assertIn(b"data-field-style-bold", certificate_page.data)
        self.assertIn(b"data-field-style-italic", certificate_page.data)
        docx_preview = self.client.post(
            "/admin/certificate-template/docx-preview",
            data={"template_docx": (io.BytesIO(starter.data), "preview.docx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(docx_preview.status_code, 200)
        self.assertGreater(docx_preview.json["token_count"], 0)
        self.assertTrue(any(
            token["token"] == "name"
            for paragraph in docx_preview.json["paragraphs"]
            for token in paragraph["tokens"]
        ))
        image_buffer = io.BytesIO()
        Image.new("RGB", (842, 595), (25, 40, 60)).save(image_buffer, format="PNG")
        image_buffer.seek(0)
        template_response = self.client.post(
            "/admin/certificates",
            data={
                "action": "template",
                "event_name": "FUTURE-TECH INNOVISION 2026",
                "event_date": "2 November 2026",
                "event_venue": "N K Orchid College",
                "type_name": "Certificate of Participation",
                "recipient_role": "team_member",
                "positions_json": json.dumps({
                    "recipient_name": {
                        "x": 61,
                        "y": 17,
                        "size": 24,
                        "font": "Helvetica",
                        "align": "center",
                        "color": "#cc1100",
                        "bold": True,
                        "italic": True,
                    },
                    "qr": {"x": 18, "y": 82, "size": 6},
                    "docx_field_styles": {
                        "name": {
                            "font": "Courier New",
                            "size": 18,
                            "color": "#aa1122",
                            "bold": True,
                            "italic": True,
                        },
                    },
                    "docx_text_mappings": [{
                        "paragraph_index": 2,
                        "text": "This certificate is presented to",
                        "field": "recipient_role",
                    }],
                    "cover_zones": [{
                        "x": 20,
                        "y": 20,
                        "width": 20,
                        "height": 10,
                        "color": "#ffffff",
                    }],
                    "custom_texts": [{
                        "id": "event-caption",
                        "text": "Awarded by {{event_name}}",
                        "x": 50,
                        "y": 50,
                        "size": 20,
                        "color": "#5326a7",
                        "font": "Times",
                        "bold": True,
                        "italic": False,
                        "align": "center",
                    }],
                }),
                "template_file": (image_buffer, "canva-template.png"),
                "template_docx": (io.BytesIO(starter.data), "editable-template.docx"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(template_response.status_code, 302)
        with self.app.app_context():
            template = CertificateTemplate.query.one()
            template_id = template.id
            self.assertFalse(template.approved)
            self.assertTrue(template.document_file.endswith(".docx"))
        saved_document = self.client.get(
            f"/admin/certificate-templates/{template_id}/document"
        )
        self.assertEqual(saved_document.status_code, 200)
        downloaded_document = saved_document.get_data()
        saved_document.close()
        self.assertEqual(downloaded_document, starter.data)
        self.client.post(f"/admin/certificate-templates/{template_id}/approve")
        incorrect_csv = io.BytesIO(
            b"Serial,Full Name,Team,Project,Domain\n"
            b"1,Alice Example,FT_01,Smart Sensor,IoT Group\n"
        )
        incorrect_mapping = self.client.post(
            "/admin/certificates",
            data={
                "action": "import_csv",
                "template_id": str(template_id),
                "username_column": "Serial",
                "recipient_name_column": "Full Name",
                "team_code_column": "Team",
                "project_title_column": "Project",
                "domain_column": "Domain",
                "recipient_csv": (incorrect_csv, "incorrect-recipients.csv"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(incorrect_mapping.status_code, 400)
        self.assertIn(b"exact login username (not a serial number", incorrect_mapping.data)
        long_recipient_name = "Alice Example " * 18
        csv_buffer = io.BytesIO(
            (
                " Student_ID , Name , Team_ID , Title , Domain , Venue , Event , Date , Role , Day \n"
                f"csv_student,{long_recipient_name},FT_01,Smart Sensor,"
                "IoT Group,Orchid Auditorium,CSV Event Override,3 November 2026,"
                "Project Leader,Monday\n"
            ).encode()
        )
        imported = self.client.post(
            "/admin/certificates",
            data={
                "action": "import_csv",
                "template_id": str(template_id),
                "recipient_csv": (csv_buffer, "recipients.csv"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(imported.status_code, 302)
        with self.app.app_context():
            certificate = Certificate.query.one()
            self.assertEqual(certificate.recipient_name, long_recipient_name.strip())
            self.assertEqual(certificate.template_id, template_id)
            certificate_fields = json.loads(certificate.fields_json)
            self.assertEqual(certificate_fields["recipient_role"], "Project Leader")
            self.assertEqual(certificate_fields["event_day"], "Monday")
            certificate_db_id = certificate.id
            certificate_id = certificate.certificate_id
            self.assertEqual(certificate.dispatch_status, "pending")
            self.assertTrue(certificate_id.startswith("FTI26-"))
            self.assertEqual(len(certificate_id), 38)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append([
            "Username", "Name", "Team_ID", "Title", "Domain", "Event", "Date",
            "Role", "Day",
            "Venue", "Email",
        ])
        worksheet.append([
            "csv_student", long_recipient_name.strip(), "FT_01", "Smart Sensor",
            "IoT Group", "CSV Event Override", "3 November 2026",
            "Project Leader", "Monday",
            "Orchid Auditorium", "csv_student@example.edu",
        ])
        excel_buffer = io.BytesIO()
        workbook.save(excel_buffer)
        excel_data = excel_buffer.getvalue()
        preview = self.client.post(
            "/admin/certificate-spreadsheet-headers",
            data={
                "recipient_csv": (io.BytesIO(excel_data), "recipients.xlsx"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(preview.status_code, 200)
        self.assertIn("Username", preview.json["headers"])
        invalid_preview = self.client.post(
            "/admin/certificate-spreadsheet-headers",
            data={
                "recipient_csv": (io.BytesIO(b"not an Excel workbook"), "invalid.xlsx"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(invalid_preview.status_code, 400)
        self.assertIn("valid .xlsx", invalid_preview.json["message"])
        with patch.dict(self.app.config, {
            "TASK_EMAIL_SMTP_HOST": "smtp.example.edu",
            "TASK_EMAIL_SMTP_PORT": 587,
            "TASK_EMAIL_SMTP_USERNAME": "mailer@example.edu",
            "TASK_EMAIL_SMTP_PASSWORD": "test-password",
            "TASK_EMAIL_FROM": "mailer@example.edu",
        }):
            excel_import = self.client.post(
                "/admin/certificates",
                data={
                    "action": "import_csv",
                    "template_id": str(template_id),
                    "email_column": "Email",
                    "recipient_csv": (
                        io.BytesIO(excel_data),
                        "recipients.xlsx",
                    ),
                },
                content_type="multipart/form-data",
            )
        self.assertEqual(excel_import.status_code, 302)
        with self.app.app_context():
            self.assertEqual(Certificate.query.count(), 1)
            certificate = Certificate.query.one()
            self.assertNotEqual(certificate.certificate_id, certificate_id)
            certificate_id = certificate.certificate_id
            self.assertEqual(certificate.dispatch_status, "pending")
        certificate_management = self.client.get("/admin/certificates")
        self.assertIn(b"Replace background (optional)", certificate_management.data)
        self.assertIn(b"Recipient email", certificate_management.data)
        self.assertIn(b"Approve and dispatch", certificate_management.data)
        self.assertIn(b".xlsx", certificate_management.data)
        self.assertIn(b"Generated certificates", certificate_management.data)
        self.assertIn(b"Download Word", certificate_management.data)
        self.client.post("/logout")
        self.assertEqual(self.login("csv_student", "Student123").status_code, 302)
        member_dashboard = self.client.get("/member")
        self.assertNotIn(certificate_id.encode(), member_dashboard.data)
        self.assertEqual(
            self.client.get(f"/certificate/{certificate_id}/download").status_code,
            403,
        )
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with patch.dict(self.app.config, {
            "TASK_EMAIL_SMTP_HOST": "smtp.example.edu",
            "TASK_EMAIL_SMTP_PORT": 587,
            "TASK_EMAIL_SMTP_USERNAME": "mailer@example.edu",
            "TASK_EMAIL_SMTP_PASSWORD": "test-password",
            "TASK_EMAIL_FROM": "mailer@example.edu",
        }):
            with patch("app.routes.smtplib.SMTP") as smtp_factory:
                smtp_server = MagicMock()
                smtp_factory.return_value.__enter__.return_value = smtp_server
                dispatched = self.client.post(
                    "/admin/certificates/dispatch",
                    data={"scope": "user", "user_id": str(member_user_id)},
                )
        self.assertEqual(dispatched.status_code, 302)
        smtp_server.send_message.assert_called_once()
        email_message = smtp_server.send_message.call_args.args[0]
        self.assertEqual(email_message["To"], "csv_student@example.edu")
        attachments = list(email_message.iter_attachments())
        self.assertEqual(len(attachments), 1)
        self.assertEqual(attachments[0].get_content_type(), "application/pdf")
        self.assertTrue(attachments[0].get_payload(decode=True).startswith(b"%PDF"))
        with self.app.app_context():
            saved_certificate = db.session.get(Certificate, certificate_db_id)
            self.assertEqual(saved_certificate.dispatch_status, "dispatched")
            self.assertIsNotNone(saved_certificate.dispatched_at)
            self.assertTrue(UserNotification.query.filter_by(
                user_id=member_user_id,
                title="Certificate dispatched: Certificate of Participation",
            ).first())
        self.client.post("/logout")
        self.assertEqual(self.login("csv_student", "Student123").status_code, 302)
        member_dashboard = self.client.get("/member")
        self.assertIn(b"Download Word", member_dashboard.data)
        pdf_response = self.client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(pdf_response.status_code, 200)
        document = pymupdf.open(stream=pdf_response.data, filetype="pdf")
        certificate_text = document[0].get_text()
        self.assertIn(certificate_id, certificate_text)
        self.assertGreaterEqual(len(document[0].get_images()), 2)
        self.assertEqual(certificate_text.count("Alice"), 18, certificate_text)
        self.assertEqual(certificate_text.count("Example"), 18, certificate_text)
        name_spans = [
            span
            for block in document[0].get_text("dict")["blocks"]
            for line in block.get("lines", [])
            for span in line["spans"]
            if "Alice" in span["text"]
        ]
        self.assertTrue(name_spans)
        self.assertTrue(all(span["color"] == 0xCC1100 for span in name_spans))
        self.assertTrue(all("Bold" in span["font"] for span in name_spans))
        self.assertTrue(all("Oblique" in span["font"] for span in name_spans))
        number_spans = [
            span
            for block in document[0].get_text("dict")["blocks"]
            for line in block.get("lines", [])
            for span in line["spans"]
            if certificate_id in span["text"]
        ]
        qr_images = [
            image["bbox"]
            for image in document[0].get_image_info()
            if (
                abs(
                    (image["bbox"][2] - image["bbox"][0])
                    - (image["bbox"][3] - image["bbox"][1])
                ) < 1
                and image["bbox"][2] - image["bbox"][0]
                < document[0].rect.width * 0.2
            )
        ]
        self.assertTrue(number_spans)
        self.assertTrue(qr_images)
        self.assertGreaterEqual(number_spans[0]["bbox"][1], qr_images[-1][3])
        self.assertIn("IoT Group", document[0].get_text())
        self.assertIn("FT_01", document[0].get_text())
        self.assertIn("Smart Sensor", document[0].get_text())
        self.assertIn("Orchid Auditorium", document[0].get_text())
        self.assertIn("CSV Event Override", document[0].get_text())
        self.assertIn("Awarded by CSV Event Override", document[0].get_text())
        self.assertIn("3 November 2026", document[0].get_text())
        rendered_page = document[0].get_pixmap()
        cover_x = int(rendered_page.width * 0.2)
        cover_y = int(rendered_page.height * 0.2)
        cover_offset = (cover_y * rendered_page.width + cover_x) * rendered_page.n
        self.assertGreaterEqual(
            min(rendered_page.samples[cover_offset:cover_offset + 3]),
            245,
        )
        document.close()
        word_response = self.client.get(
            f"/certificate/{certificate_id}/word-download"
        )
        self.assertEqual(word_response.status_code, 200)
        merged_docx = zipfile.ZipFile(io.BytesIO(word_response.data))
        merged_xml = merged_docx.read("word/document.xml")
        self.assertIn(b"Courier New", merged_xml)
        self.assertIn(b"AA1122", merged_xml)
        self.assertIn(b"Project Leader", merged_xml)
        self.assertNotIn(b"This certificate is presented to", merged_xml)
        self.assertIn(b"<w:b w:val=\"1\"", merged_xml)
        self.assertIn(b"<w:i w:val=\"1\"", merged_xml)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        drawing_ns = (
            "http://schemas.openxmlformats.org/drawingml/2006/"
            "wordprocessingDrawing"
        )
        merged_root = ET.fromstring(merged_xml)
        anchors = merged_root.findall(f".//{{{drawing_ns}}}anchor")
        self.assertTrue(anchors)
        position_h = anchors[-1].find(f"{{{drawing_ns}}}positionH")
        position_v = anchors[-1].find(f"{{{drawing_ns}}}positionV")
        extent = anchors[-1].find(f"{{{drawing_ns}}}extent")
        page_size = merged_root.find(f".//{{{word_ns}}}sectPr/{{{word_ns}}}pgSz")
        self.assertIsNotNone(position_h)
        self.assertIsNotNone(position_v)
        self.assertIsNotNone(extent)
        self.assertIsNotNone(page_size)
        page_width_emu = int(page_size.get(f"{{{word_ns}}}w")) * 635
        page_height_emu = int(page_size.get(f"{{{word_ns}}}h")) * 635
        expected_qr_size = int(page_width_emu * 6 / 100)
        expected_qr_left = int(page_width_emu * 18 / 100 - int(extent.get("cx")) / 2)
        expected_qr_top = int(page_height_emu * 82 / 100 - expected_qr_size / 2)
        self.assertAlmostEqual(
            int(position_h.findtext(f"{{{drawing_ns}}}posOffset")),
            expected_qr_left,
            delta=2,
        )
        self.assertAlmostEqual(
            int(position_v.findtext(f"{{{drawing_ns}}}posOffset")),
            expected_qr_top,
            delta=2,
        )
        self.assertGreater(int(extent.get("cx")), expected_qr_size)
        self.assertGreater(int(extent.get("cy")), expected_qr_size)
        self.assertIn(long_recipient_name.strip().encode(), merged_xml)
        self.assertNotIn(certificate_id.encode(), merged_xml)
        self.assertNotIn(b"&lt;name&gt;", merged_xml)
        self.assertNotIn(b"&lt;qr&gt;", merged_xml)
        qr_image_names = [
            name for name in merged_docx.namelist()
            if name.startswith("word/media/certificate-qr-")
        ]
        self.assertEqual(len(qr_image_names), 1)
        merged_qr = Image.open(io.BytesIO(merged_docx.read(qr_image_names[0])))
        self.assertGreater(merged_qr.width, merged_qr.height)
        self.assertGreater(
            sum(
                1
                for pixel in merged_qr.crop(
                    (0, merged_qr.width // 3, merged_qr.width, merged_qr.height)
                ).convert("RGB").getdata()
                if max(pixel) < 100
            ),
            0,
        )
        merged_docx.close()
        self.client.post("/logout")
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        updated_template = self.client.post(
            f"/admin/certificate-templates/{template_id}/edit",
            data={
                "event_name": "Updated Event Design",
                "event_date": "4 November 2026",
                "event_venue": "Updated Venue",
                "recipient_role": "team_member",
                "positions_json": "{}",
            },
        )
        self.assertEqual(updated_template.status_code, 302)
        with self.app.app_context():
            saved_certificate = db.session.get(Certificate, certificate_db_id)
            snapshot = json.loads(saved_certificate.template_snapshot_json)
            self.assertEqual(snapshot["event_name"], "FUTURE-TECH INNOVISION 2026")
            self.assertTrue(snapshot["document_file"])
            self.assertTrue(
                os.path.isfile(
                    os.path.join(
                        self.upload_dir.name,
                        "certificate_templates",
                        snapshot["background_file"],
                    )
                )
            )
            self.assertTrue(
                os.path.isfile(
                    os.path.join(
                        self.upload_dir.name,
                        "certificate_templates",
                        snapshot["document_file"],
                    )
                )
            )
        old_design_download = self.client.get(
            f"/certificate/{certificate_id}/download"
        )
        self.assertEqual(old_design_download.status_code, 200)
        old_design_pdf = pymupdf.open(
            stream=old_design_download.data,
            filetype="pdf",
        )
        self.assertIn("CSV Event Override", old_design_pdf[0].get_text())
        old_design_pdf.close()
        old_design_word = self.client.get(
            f"/certificate/{certificate_id}/word-download"
        )
        self.assertEqual(old_design_word.status_code, 200)
        self.assertIn(
            long_recipient_name.strip().encode(),
            zipfile.ZipFile(io.BytesIO(old_design_word.data)).read(
                "word/document.xml"
            ),
        )
        self.assertEqual(
            self.client.post(
                f"/admin/certificate-templates/{template_id}/delete"
            ).status_code,
            302,
        )
        preserved_word = self.client.get(
            f"/certificate/{certificate_id}/word-download"
        )
        self.assertEqual(preserved_word.status_code, 200)
        self.client.post(f"/admin/certificates/{certificate_db_id}/delete")
        with self.app.app_context():
            self.assertIsNone(db.session.get(Certificate, certificate_db_id))

    def test_word_only_certificate_template_creates_pdf_and_renders_certificate(self):
            self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
            starter = self.client.get(
                "/admin/certificate-template-starter"
                "?type_name=Certificate%20of%20Participation"
            )
            self.assertEqual(starter.status_code, 200)

            pdf_document = pymupdf.open()
            pdf_document.new_page(width=842, height=595)
            rendered_pdf = pdf_document.tobytes()
            pdf_document.close()

            with patch(
                "app.routes.convert_certificate_docx_to_pdf",
                return_value=rendered_pdf,
            ) as convert_docx:
                preview = self.client.post(
                    "/admin/certificate-template/docx-preview-pdf",
                    data={
                        "template_docx": (
                            io.BytesIO(starter.data),
                            "preview.docx",
                        ),
                    },
                    content_type="multipart/form-data",
                )
                self.assertEqual(preview.status_code, 200)
                self.assertEqual(preview.mimetype, "application/pdf")
                self.assertTrue(preview.data.startswith(b"%PDF"))
                self.assertEqual(
                    convert_docx.call_args.args[0],
                    starter.data,
                )

                created = self.client.post(
                    "/admin/certificates",
                    data={
                        "action": "template",
                        "event_name": "FUTURE-TECH INNOVISION 2026",
                        "event_date": "7 November 2026",
                        "event_venue": "N K Orchid College",
                        "type_name": "Certificate of Participation",
                        "recipient_role": "team_member",
                        "positions_json": "{}",
                        "template_docx": (
                            io.BytesIO(starter.data),
                            "word-only.docx",
                        ),
                    },
                    content_type="multipart/form-data",
                )
                self.assertEqual(created.status_code, 302, created.get_data(as_text=True))

                with self.app.app_context():
                    template = CertificateTemplate.query.one()
                    self.assertTrue(template.background_file.endswith(".pdf"))
                    self.assertTrue(template.document_file.endswith(".docx"))
                    positions = json.loads(template.positions_json)
                    self.assertTrue(positions["docx_pdf_primary"])
                    template_id = template.id
                    certificate_type_id = template.type_id
                    certificate = Certificate(
                        certificate_id="FTI26-WORDONLY-TEST",
                        type_id=certificate_type_id,
                        template_id=template_id,
                        recipient_name="Word Only Recipient",
                        recipient_role="Team Member",
                        team_code="FT_01",
                        fields_json=json.dumps({
                            "event_name": "FUTURE-TECH INNOVISION 2026",
                        }),
                        dispatch_status="dispatched",
                        template_snapshot_json=json.dumps({
                            "background_file": template.background_file,
                            "document_file": template.document_file,
                            "positions_json": template.positions_json,
                            "page_width": template.page_width,
                            "page_height": template.page_height,
                        }),
                    )
                    db.session.add(certificate)
                    db.session.commit()

                download = self.client.get(
                    "/certificate/FTI26-WORDONLY-TEST/download"
                )
                self.assertEqual(download.status_code, 200)
                self.assertEqual(download.mimetype, "application/pdf")
                self.assertEqual(download.data, rendered_pdf)

    def test_merged_word_certificate_keeps_opc_default_namespaces(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        starter = self.client.get(
            "/admin/certificate-template-starter"
            "?type_name=Certificate%20of%20Participation"
        )
        self.assertEqual(starter.status_code, 200)
        merged = merged_certificate_docx(
            io.BytesIO(starter.data),
            {
                "recipient_name": "Test Recipient",
                "certificate_id": "FTI26-MERGE-TEST",
            },
            "https://example.test/certificates/FTI26-MERGE-TEST",
        )
        with zipfile.ZipFile(merged) as archive:
            relationships = archive.read("word/_rels/document.xml.rels")
            content_types = archive.read("[Content_Types].xml")
        self.assertIn(
            b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"',
            relationships,
        )
        self.assertIn(
            b'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"',
            content_types,
        )
        self.assertNotIn(b"ns0:", relationships)
        self.assertNotIn(b"ns0:", content_types)

    def test_onlyoffice_editor_saves_word_and_invalidates_approval(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        starter = self.client.get(
            "/admin/certificate-template-starter"
            "?type_name=Certificate%20of%20Participation"
        )
        self.assertEqual(starter.status_code, 200)
        pdf_document = pymupdf.open()
        pdf_document.new_page(width=842, height=595)
        rendered_pdf = pdf_document.tobytes()
        pdf_document.close()
        with patch(
            "app.routes.convert_certificate_docx_to_pdf",
            return_value=rendered_pdf,
        ):
            created = self.client.post(
                "/admin/certificates",
                data={
                    "action": "template",
                    "event_name": "FUTURE-TECH INNOVISION 2026",
                    "event_date": "7 November 2026",
                    "event_venue": "N K Orchid College",
                    "type_name": "Certificate of Participation",
                    "recipient_role": "team_member",
                    "positions_json": "{}",
                    "template_docx": (
                        io.BytesIO(starter.data),
                        "onlyoffice-word.docx",
                    ),
                },
                content_type="multipart/form-data",
            )
        self.assertEqual(created.status_code, 302)
        with self.app.app_context():
            template = CertificateTemplate.query.one()
            template_id = template.id
            old_document_file = template.document_file
            old_background_file = template.background_file
            db.session.add(Certificate(
                certificate_id="FTI26-ONLYOFFICE-HISTORY",
                type_id=template.type_id,
                template_id=template.id,
                recipient_name="Historical Recipient",
                recipient_role="Team Member",
                dispatch_status="dispatched",
            ))
            db.session.commit()
        self.assertEqual(
            self.client.post(
                f"/admin/certificate-templates/{template_id}/approve"
            ).status_code,
            302,
        )
        certificate_page = self.client.get("/admin/certificates")
        self.assertEqual(certificate_page.status_code, 200)
        self.assertIn(
            b"The in-page Word editor will be available",
            certificate_page.data,
        )

        previous_server_url = self.app.config.get("ONLYOFFICE_SERVER_URL", "")
        previous_secret = self.app.config.get("ONLYOFFICE_JWT_SECRET", "")
        self.app.config.update(
            ONLYOFFICE_SERVER_URL="https://docs.example.test",
            ONLYOFFICE_JWT_SECRET="a-test-onlyoffice-secret",
        )
        try:
            config_response = self.client.get(
                f"/admin/certificate-templates/{template_id}/editor-config"
            )
            self.assertEqual(config_response.status_code, 200)
            editor = config_response.get_json()
            editor_config = editor["config"]
            token_payload = decode_onlyoffice_token(
                editor_config["token"], "a-test-onlyoffice-secret"
            )
            self.assertEqual(token_payload["document"]["key"], editor_config["document"]["key"])
            self.assertEqual(
                editor_config["document"]["url"].split("/", 3)[2],
                "localhost",
            )
            signed_document = self.client.get(
                editor_config["document"]["url"].split("localhost", 1)[1]
            )
            self.assertEqual(signed_document.status_code, 200)
            self.assertTrue(signed_document.data.startswith(b"PK"))
            signed_document.close()

            callback_body = {
                "key": editor_config["document"]["key"],
                "status": 2,
                "url": "https://docs.example.test/edited-certificate.docx",
            }
            callback_token = encode_onlyoffice_token(
                callback_body, "a-test-onlyoffice-secret"
            )
            rejected_callback = self.client.post(
                f"/onlyoffice/certificate-templates/{template_id}/callback"
                f"?key={callback_body['key']}",
                json=callback_body,
                headers={"Authorization": "Bearer invalid-token"},
            )
            self.assertEqual(rejected_callback.status_code, 403)
            download = MagicMock()
            download.__enter__.return_value = download
            download.geturl.return_value = callback_body["url"]
            download.read.return_value = starter.data
            office_opener = MagicMock()
            office_opener.open.return_value = download
            with (
                patch(
                    "app.routes.urllib.request.build_opener",
                    return_value=office_opener,
                ),
                patch(
                    "app.routes.convert_certificate_docx_to_pdf",
                    return_value=rendered_pdf,
                ),
            ):
                callback_response = self.client.post(
                    f"/onlyoffice/certificate-templates/{template_id}/callback"
                    f"?key={callback_body['key']}",
                    json=callback_body,
                    headers={"Authorization": f"Bearer {callback_token}"},
                )
            self.assertEqual(callback_response.status_code, 200)
            self.assertEqual(callback_response.get_json(), {"error": 0})
            with self.app.app_context():
                template = db.session.get(CertificateTemplate, template_id)
                self.assertFalse(template.approved)
                self.assertNotEqual(template.document_file, editor["revision"])
                historical = Certificate.query.filter_by(
                    certificate_id="FTI26-ONLYOFFICE-HISTORY"
                ).one()
                snapshot = json.loads(historical.template_snapshot_json)
                self.assertEqual(snapshot["document_file"], old_document_file)
                self.assertEqual(snapshot["background_file"], old_background_file)
                self.assertTrue(
                    os.path.isfile(
                        os.path.join(
                            self.app.config["UPLOAD_FOLDER"],
                            "certificate_templates",
                            old_document_file,
                        )
                    )
                )
                self.assertTrue(
                    os.path.isfile(
                        os.path.join(
                            self.app.config["UPLOAD_FOLDER"],
                            "certificate_templates",
                            template.document_file,
                        )
                    )
                )
        finally:
            self.app.config.update(
                ONLYOFFICE_SERVER_URL=previous_server_url,
                ONLYOFFICE_JWT_SECRET=previous_secret,
            )

    def test_deleting_template_preserves_issued_certificate_download(self):
        image_buffer = io.BytesIO()
        Image.new("RGB", (524, 724), "white").save(image_buffer, format="PNG")
        background_filename = "delete-template-preserved.png"
        template_directory = os.path.join(
            self.upload_dir.name, "certificate_templates"
        )
        os.makedirs(template_directory, exist_ok=True)
        with open(os.path.join(template_directory, background_filename), "wb") as file:
            file.write(image_buffer.getvalue())

        with self.app.app_context():
            certificate_type = CertificateType.query.filter_by(
                name="Certificate of Participation"
            ).one()
            template = CertificateTemplate(
                type_id=certificate_type.id,
                event_name="FUTURE-TECH INNOVISION 2026",
                event_date="2 November 2026",
                event_venue="Orchid Auditorium",
                recipient_role="team_member",
                background_file=background_filename,
                positions_json=json.dumps({
                    "custom_texts": [{
                        "id": "award-line",
                        "text": "Presented at {{event_name}}",
                        "x": 50,
                        "y": 35,
                        "size": 20,
                        "color": "#20242c",
                        "font": "Helvetica",
                        "bold": True,
                        "italic": False,
                        "align": "center",
                    }],
                }),
                page_width=524,
                page_height=724,
                approved=True,
            )
            db.session.add(template)
            db.session.flush()
            template_id = template.id
            certificate = Certificate(
                certificate_id="FTI26-TEMPLATEDELETE",
                type_id=certificate_type.id,
                recipient_name="Preserved Recipient",
                recipient_role="Team Member",
                template_id=template.id,
                fields_json=json.dumps({"domain_name": "CodeWave"}),
            )
            db.session.add(certificate)
            db.session.commit()
            certificate_id = certificate.certificate_id

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        page = self.client.get("/admin/certificates")
        self.assertIn(b"Delete template", page.data)
        self.assertEqual(
            self.client.post(
                f"/admin/certificate-templates/{template_id}/delete"
            ).status_code,
            302,
        )
        with self.app.app_context():
            self.assertIsNone(db.session.get(CertificateTemplate, template_id))
            preserved = Certificate.query.filter_by(
                certificate_id=certificate_id
            ).one()
            self.assertIsNone(preserved.template_id)
            snapshot = json.loads(preserved.template_snapshot_json)
            self.assertEqual(snapshot["background_file"], background_filename)
            self.assertTrue(
                os.path.isfile(os.path.join(template_directory, background_filename))
            )
        download = self.client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(download.status_code, 200)
        document = pymupdf.open(stream=download.data, filetype="pdf")
        self.assertIn("Presented at FUTURE-TECH INNOVISION 2026", document[0].get_text())
        document.close()

    def test_bulk_certificate_deletion_allows_template_editing_again(self):
        with self.app.app_context():
            certificate_type = CertificateType.query.filter_by(
                name="Certificate of Appreciation - Mentor"
            ).one()
            template = CertificateTemplate(
                type_id=certificate_type.id,
                event_name="FUTURE-TECH INNOVISION 2026",
                event_date="2 November 2026",
                event_venue="Orchid Auditorium",
                recipient_role="mentor",
                background_file="bulk-template.png",
                positions_json="{}",
                approved=True,
            )
            db.session.add(template)
            db.session.flush()
            certificate = Certificate(
                certificate_id="FTI26-BULKDELETE",
                type_id=certificate_type.id,
                recipient_name="Delete Me",
                recipient_role="Team Member",
                template_id=template.id,
                file_path="bulk-certificate.pdf",
            )
            db.session.add(certificate)
            db.session.commit()
            template_id = template.id
        certificate_path = os.path.join(
            self.upload_dir.name, "certificates", "bulk-certificate.pdf"
        )
        with open(certificate_path, "wb") as file:
            file.write(b"%PDF test")

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        self.assertEqual(
            self.client.post("/admin/certificates/delete-all").status_code,
            302,
        )
        with self.app.app_context():
            self.assertEqual(Certificate.query.count(), 0)
            self.assertIsNotNone(db.session.get(CertificateTemplate, template_id))
        self.assertFalse(os.path.exists(certificate_path))
        edited = self.client.post(
            f"/admin/certificate-templates/{template_id}/edit",
            data={
                "event_name": "Updated Event",
                "event_date": "3 November 2026",
                "event_venue": "Updated Venue",
                "recipient_role": "mentor",
                "positions_json": "{}",
            },
        )
        self.assertEqual(edited.status_code, 302)

    def test_portrait_role_certificate_keeps_reference_page_shape(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            account = User(
                username="portrait_member",
                role="team_member",
                team_id=team.id,
            )
            account.set_password("PortraitPass12")
            db.session.add(account)
            db.session.commit()

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        portrait = Image.new("RGB", (524, 724), "white")
        image_buffer = io.BytesIO()
        portrait.save(image_buffer, format="PNG")
        image_buffer.seek(0)
        saved = self.client.post(
            "/admin/certificates",
            data={
                "action": "template",
                "event_name": "FUTURE-TECH INNOVISION 2026",
                "event_date": "2 November 2026",
                "event_venue": "Orchid Auditorium",
                "type_name": "Certificate of Participation",
                "recipient_role": "team_member",
                "positions_json": "{}",
                "template_file": (image_buffer, "portrait-template.png"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(saved.status_code, 302)
        with self.app.app_context():
            template = CertificateTemplate.query.one()
            self.assertEqual(template.recipient_role, "team_member")
            self.assertAlmostEqual(template.page_width / template.page_height, 524 / 724)
            template_id = template.id
        certificate_page = self.client.get("/admin/certificates")
        self.assertEqual(certificate_page.status_code, 200)
        self.assertIn(b"Certificate studio", certificate_page.data)
        self.assertIn(b"recipient_name", certificate_page.data)
        self.assertIn(b"certificate_number", certificate_page.data)
        self.client.post(f"/admin/certificate-templates/{template_id}/approve")
        csv_buffer = io.BytesIO(
            b"Username,Full Name,Team,Project,Domain\n"
            b"portrait_member,Portrait Student,FT_01,Smart Sensor,"
            b"IoT & Robotics / Multimedia Group\n"
        )
        issued = self.client.post(
            "/admin/certificates",
            data={
                "action": "import_csv",
                "template_id": str(template_id),
                "username_column": "Username",
                "recipient_name_column": "Full Name",
                "team_code_column": "Team",
                "project_title_column": "Project",
                "domain_column": "Domain",
                "recipient_csv": (csv_buffer, "portrait-recipients.csv"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(issued.status_code, 302)
        with self.app.app_context():
            certificate = Certificate.query.one()
            certificate_id = certificate.certificate_id
            account_id = certificate.user_id
        self.assertEqual(self.client.post(
            "/admin/certificates/dispatch",
            data={"scope": "user", "user_id": str(account_id)},
        ).status_code, 302)
        self.client.post("/logout")
        self.assertEqual(self.login("portrait_member", "PortraitPass12").status_code, 302)
        pdf = self.client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(pdf.status_code, 200)
        document = pymupdf.open(stream=pdf.data, filetype="pdf")
        page = document[0].rect
        self.assertAlmostEqual(page.width / page.height, 524 / 724, places=2)
        self.assertIn("Portrait Student", document[0].get_text())
        document.close()

    def test_post_forms_reject_requests_without_csrf_tokens(self):
        self.app.config["WTF_CSRF_ENABLED"] = True
        try:
            response = self.client.post(
                "/login",
                data={"username": "admin", "password": "Admin@12345!"},
            )
            self.assertEqual(response.status_code, 400)
        finally:
            self.app.config["WTF_CSRF_ENABLED"] = False

    def test_admin_workflow_templates_render(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            team_id = Team.query.filter_by(team_code="FT_01").one().id
            judge = User(
                username="assignment_page_judge",
                role="judge",
                display_name="Assignment Page Judge",
            )
            judge.set_password("AssignmentPagePassword123!")
            db.session.add(judge)
            db.session.flush()
            db.session.add(JudgeAssignment(judge_id=judge.id, team_id=team_id))
            db.session.commit()
        assignment_page = self.client.get("/admin/judge-assignments")
        self.assertEqual(assignment_page.status_code, 200)
        self.assertIn(b"Assignment Page Judge", assignment_page.data)
        self.assertIn(
            f'value="{team_id}" checked'.encode(),
            assignment_page.data,
        )
        for path in (
            "/admin",
            "/admin/teams",
            "/admin/users",
            "/admin/judge-assignments",
            "/admin/role-assignments",
            "/admin/branding",
            "/admin/monitor",
            "/admin/announcements",
            f"/admin/admin-evaluation/{team_id}",
            "/account/password",
        ):
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 200)
                if path == "/admin/teams":
                    self.assertIn(b"Project Management", response.data)
                    self.assertIn(b"admin-team-table", response.data)
                if path == "/admin/users":
                    self.assertIn(b"Select one or two teams for this coordinator.", response.data)

    def test_default_branding_assets_render_without_uploaded_files(self):
        with self.app.app_context():
            db.session.add_all(
                [
                    SiteSetting(key="college_logo", value="college_logo.jpg"),
                    SiteSetting(key="club_logo", value="club_logo.png"),
                    SiteSetting(
                        key="college_banner_image", value="old_uploaded_header.png"
                    ),
                ]
            )
            db.session.commit()

        response = self.client.get("/login")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"/branding/college_logo.png", response.data)
        self.assertIn(b"/branding/club_logo.jpg", response.data)
        self.assertIn(b"/branding/header.png", response.data)
        self.assertIn(b"/static/branding/favicon.png?v=1", response.data)

        for filename in ("college_logo.png", "club_logo.jpg", "header.png"):
            with self.subTest(filename=filename):
                asset = self.client.get(f"/branding/{filename}")
                self.assertEqual(asset.status_code, 200)
                self.assertGreater(len(asset.data), 0)
                asset.close()

    def test_coordinator_can_only_mark_attendance_for_assigned_teams(self):
        with self.app.app_context():
            teams = Team.query.filter(Team.team_code.in_(["FT_01", "FT_02"])).order_by(Team.team_code).all()
            first_member = TeamMember(team_id=teams[0].id, full_name="Assigned Student", prn="ATTEND001")
            second_member = TeamMember(team_id=teams[1].id, full_name="Other Student", prn="ATTEND002")
            coordinator = User(username="attendance_coordinator", role="coordinator")
            coordinator.set_password("AttendancePassword123!")
            db.session.add_all([first_member, second_member, coordinator])
            db.session.flush()
            db.session.add(CoordinatorTeamAssignment(
                coordinator_id=coordinator.id,
                team_id=teams[0].id,
            ))
            db.session.commit()
            first_member_id, second_member_id = first_member.id, second_member.id

        self.assertEqual(self.login("attendance_coordinator", "AttendancePassword123!").status_code, 302)
        roster = self.client.get("/attendance?event_date=2026-03-10")
        self.assertEqual(roster.status_code, 200)
        self.assertIn(b"Assigned Student", roster.data)
        self.assertNotIn(b"Other Student", roster.data)
        response = self.client.post("/attendance", data={
            "event_date": "2026-03-10",
            f"status_{first_member_id}": "present",
        })
        self.assertEqual(response.status_code, 302)
        denied = self.client.post("/attendance", data={
            "event_date": "2026-03-10",
            f"status_{second_member_id}": "present",
        })
        self.assertEqual(denied.status_code, 403)
        with self.app.app_context():
            record = Attendance.query.one()
            self.assertEqual(record.member_id, first_member_id)
            self.assertEqual(record.status, "present")

    def test_judge_can_only_mark_attendance_for_assigned_teams(self):
        with self.app.app_context():
            teams = Team.query.filter(
                Team.team_code.in_(["FT_01", "FT_02"])
            ).order_by(Team.team_code).all()
            assigned_member = TeamMember(
                team_id=teams[0].id,
                full_name="Judge Assigned Student",
                prn="JUDGEATTEND001",
            )
            other_member = TeamMember(
                team_id=teams[1].id,
                full_name="Judge Other Student",
                prn="JUDGEATTEND002",
            )
            judge = User(username="attendance_judge", role="judge")
            judge.set_password("JudgeAttendance1")
            db.session.add_all([assigned_member, other_member, judge])
            db.session.flush()
            db.session.add(
                JudgeAssignment(judge_id=judge.id, team_id=teams[0].id)
            )
            db.session.commit()
            assigned_member_id, other_member_id = assigned_member.id, other_member.id

        self.assertEqual(
            self.login("attendance_judge", "JudgeAttendance1").status_code, 302
        )
        dashboard = self.client.get("/judge")
        self.assertEqual(dashboard.status_code, 200)
        self.assertIn(b"Take attendance", dashboard.data)
        roster = self.client.get("/attendance?event_date=2026-03-11")
        self.assertEqual(roster.status_code, 200)
        self.assertIn(b"Judge Assigned Student", roster.data)
        self.assertNotIn(b"Judge Other Student", roster.data)
        accepted = self.client.post(
            "/attendance",
            data={
                "event_date": "2026-03-11",
                f"status_{assigned_member_id}": "present",
            },
        )
        self.assertEqual(accepted.status_code, 302)
        denied = self.client.post(
            "/attendance",
            data={
                "event_date": "2026-03-11",
                f"status_{other_member_id}": "present",
            },
        )
        self.assertEqual(denied.status_code, 403)

    def test_admin_can_delete_announcement_and_account(self):
        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        with self.app.app_context():
            account = User(username="remove_me", role="mentor", display_name="Remove Me")
            account.set_password("DeletePass1")
            announcement = Announcement(title="Temporary update", body="Remove this update")
            db.session.add_all([account, announcement])
            db.session.commit()
            account_id, announcement_id = account.id, announcement.id

        accounts_page = self.client.get("/admin/users")
        self.assertEqual(accounts_page.status_code, 200)
        self.assertIn(b"Delete", accounts_page.data)
        self.assertEqual(
            self.client.post(f"/admin/users/{account_id}/delete").status_code, 302
        )
        self.assertEqual(
            self.client.post(
                f"/admin/announcements/{announcement_id}/delete"
            ).status_code,
            302,
        )
        with self.app.app_context():
            self.assertIsNone(db.session.get(User, account_id))
            self.assertIsNone(db.session.get(Announcement, announcement_id))

    def test_admin_announcement_sends_in_app_notification_to_active_users(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            recipient = User(
                username="broadcast_member",
                role="team_member",
                display_name="Broadcast Member",
                team_id=team.id,
            )
            recipient.set_password("BroadcastPassword123!")
            db.session.add(recipient)
            db.session.flush()
            recipient_id = recipient.id
            inactive_recipient = User(
                username="inactive_broadcast_member",
                role="team_member",
                active=False,
                team_id=team.id,
            )
            inactive_recipient.set_password("InactiveBroadcastPassword123!")
            db.session.add(inactive_recipient)
            db.session.commit()

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        published = self.client.post("/admin/announcements", data={
            "title": "Club-wide notification test",
            "body": "Please review this update in your account.",
            "notify_all_active": "1",
        })
        self.assertEqual(published.status_code, 200)
        with self.app.app_context():
            notification = UserNotification.query.filter_by(
                user_id=recipient_id,
                title="Club-wide notification test",
            ).one()
            notification_id = notification.id
            self.assertFalse(notification.read)
            self.assertFalse(UserNotification.query.join(
                User, UserNotification.user_id == User.id
            ).filter(
                User.username == "inactive_broadcast_member",
                UserNotification.title == "Club-wide notification test",
            ).first())

        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "broadcast_member",
            "password": "BroadcastPassword123!",
            "role": "team_member",
        }).status_code, 302)
        page = member_client.get("/")
        self.assertIn(b"Club-wide notification test", page.data)
        self.assertIn(b"Please review this update in your account.", page.data)
        self.assertIn(b"Mark as read", page.data)
        self.assertEqual(self.client.post(
            f"/account/notifications/{notification_id}/read"
        ).status_code, 404)
        with self.app.app_context():
            self.assertFalse(db.session.get(UserNotification, notification_id).read)
        marked = member_client.post(
            f"/account/notifications/{notification_id}/read",
            headers={"Referer": "/"},
        )
        self.assertEqual(marked.status_code, 302)
        updated_page = member_client.get("/")
        self.assertNotIn(b"Mark as read", updated_page.data)
        with self.app.app_context():
            self.assertTrue(db.session.get(UserNotification, notification_id).read)

    def test_eligible_certificates_render_pdf_and_are_recipient_scoped(self):
        image_buffer = io.BytesIO()
        Image.new("RGB", (524, 724), "white").save(image_buffer, format="PNG")
        background_filename = "eligible-certificate-template.png"
        template_directory = os.path.join(
            self.upload_dir.name, "certificate_templates"
        )
        os.makedirs(template_directory, exist_ok=True)
        with open(os.path.join(template_directory, background_filename), "wb") as file:
            file.write(image_buffer.getvalue())
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            team.domain_id = Domain.query.order_by(Domain.id).first().id
            domain_id = team.domain_id
            account = User(username="certificate_member", role="team_member", team_id=team.id)
            account.set_password("CertificatePassword123!")
            member = TeamMember(
                team_id=team.id,
                full_name="Certificate Student",
                prn="CERT001",
                user=account,
            )
            db.session.add(member)
            db.session.commit()
            account_id = account.id
            certificate_type = CertificateType.query.filter_by(
                name="Certificate of Participation"
            ).one()
            certificate_type_id = certificate_type.id
            for certificate_type_name in (
                "Certificate of Participation",
                "Certificate of Winner",
                "Certificate of Appreciation - Mentor",
                "Certificate of Appreciation - Coordinator",
            ):
                template_type = CertificateType.query.filter_by(
                    name=certificate_type_name
                ).one()
                recipient_role = (
                    "team_member"
                    if certificate_type_name in {
                        "Certificate of Participation",
                        "Certificate of Winner",
                    }
                    else "mentor"
                    if "Mentor" in certificate_type_name
                    else "coordinator"
                )
                db.session.add(CertificateTemplate(
                    type_id=template_type.id,
                    event_name="FUTURE-TECH INNOVISION 2026",
                    event_date="2 November 2026",
                    event_venue="Orchid Auditorium",
                    recipient_role=recipient_role,
                    background_file=background_filename,
                    positions_json="{}",
                    page_width=524,
                    page_height=724,
                    approved=True,
                ))
            db.session.commit()

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        issued = self.client.post("/admin/certificates", data={
            "action": "eligible",
            "type_id": certificate_type_id,
        })
        self.assertEqual(issued.status_code, 302)
        with self.app.app_context():
            certificate = Certificate.query.filter_by(user_id=account_id).one()
            certificate_id = certificate.certificate_id
            layout_certificate_ids = []
            for certificate_type_name, suffix in (
                ("Certificate of Winner", "WINNER"),
                ("Certificate of Appreciation - Mentor", "MENTOR"),
                ("Certificate of Appreciation - Coordinator", "COORD"),
            ):
                certificate_type = CertificateType.query.filter_by(
                    name=certificate_type_name
                ).one()
                template_id = CertificateTemplate.query.filter_by(
                    type_id=certificate_type.id
                ).one().id
                layout_id = f"FTI26-LAYOUT-{suffix}"
                db.session.add(Certificate(
                    certificate_id=layout_id,
                    type_id=certificate_type.id,
                    recipient_name="Certificate Student",
                    recipient_role="Team Member",
                    team_code="FT_01",
                    user_id=account_id,
                    template_id=template_id,
                ))
                layout_certificate_ids.append(layout_id)
            db.session.commit()

        self.assertEqual(self.client.post(
            "/admin/certificates/dispatch",
            data={"scope": "domain", "domain_id": str(domain_id)},
        ).status_code, 302)
        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "certificate_member",
            "password": "CertificatePassword123!",
            "role": "team_member",
        }).status_code, 302)
        pdf = member_client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(pdf.status_code, 200)
        self.assertEqual(pdf.mimetype, "application/pdf")
        self.assertTrue(pdf.data.startswith(b"%PDF"))
        generated_pdf = pymupdf.open(stream=pdf.data, filetype="pdf")
        self.assertEqual(generated_pdf.page_count, 1)
        certificate_text = " ".join(generated_pdf[0].get_text().split())
        self.assertIn("Certificate Student", certificate_text)
        self.assertIn(certificate_id, certificate_text)
        self.assertNotIn("Pratiksha Prakash Paul", certificate_text)
        self.assertIn("2 November 2026", certificate_text)
        for layout_id in layout_certificate_ids:
            response = member_client.get(f"/certificate/{layout_id}/download")
            self.assertEqual(response.status_code, 200)
            layout_pdf = pymupdf.open(stream=response.data, filetype="pdf")
            self.assertEqual(layout_pdf.page_count, 1)
            self.assertIn("Certificate Student", layout_pdf[0].get_text())
        other_client = self.app.test_client()
        self.assertEqual(other_client.post("/login", data={
            "username": "FT_02",
            "password": "ChangeMe123!",
            "role": "team_leader",
        }).status_code, 302)
        self.assertEqual(
            other_client.get(f"/certificate/{certificate_id}/download").status_code,
            403,
        )

    def test_admin_certificate_zip_matches_student_ids_and_scopes_downloads(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            account = User(
                username="zip_certificate_member",
                role="team_member",
                display_name="Zip Student",
                team_id=team.id,
            )
            account.set_password("ZipCertificatePassword123!")
            member = TeamMember(
                team_id=team.id,
                full_name="Zip Student",
                prn="ZIPSTUDENT001",
                user=account,
            )
            db.session.add(member)
            leader_team = Team.query.filter_by(team_code="FT_02").one()
            leader_account = User(
                username="zip_certificate_leader",
                role="team_leader",
                display_name="Zip Team Leader",
                team_id=leader_team.id,
            )
            leader_account.set_password("ZipLeaderPassword123!")
            db.session.add(TeamMember(
                team_id=leader_team.id,
                full_name="Zip Team Leader",
                prn="ZIPLEADER001",
                user=leader_account,
            ))
            db.session.commit()
            account_id = account.id
            leader_account_id = leader_account.id
            type_id = CertificateType.query.filter_by(
                name="Certificate of Participation"
            ).one().id

        def make_pdf(text):
            document = pymupdf.open()
            page = document.new_page()
            page.insert_text((72, 72), text)
            return document.tobytes()

        archive_buffer = io.BytesIO()
        pdf_bytes = make_pdf("Student certificate")
        leader_pdf_bytes = make_pdf("Team leader certificate")
        with zipfile.ZipFile(archive_buffer, "w") as archive:
            archive.writestr("ZIPLEADER001.pdf", leader_pdf_bytes)
            archive.writestr("UNKNOWN_STUDENT.pdf", make_pdf("Unmatched"))
        archive_buffer.seek(0)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        response = self.client.post("/admin/certificates", data={
            "action": "upload_files",
            "type_id": str(type_id),
            "certificate_files": [
                (io.BytesIO(pdf_bytes), "ZIPSTUDENT001.pdf"),
                (archive_buffer, "student_certificates.zip"),
            ],
        }, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            certificate = Certificate.query.filter_by(user_id=account_id).one()
            certificate_id = certificate.certificate_id
            self.assertTrue(certificate.file_path)
            leader_certificate = Certificate.query.filter_by(
                user_id=leader_account_id
            ).one()
            leader_certificate_id = leader_certificate.certificate_id

        self.assertEqual(self.client.post(
            "/admin/certificates/dispatch",
            data={"scope": "all"},
        ).status_code, 302)
        member_client = self.app.test_client()
        self.assertEqual(member_client.post("/login", data={
            "username": "zip_certificate_member",
            "password": "ZipCertificatePassword123!",
            "role": "team_member",
        }).status_code, 302)
        member_page = member_client.get("/member")
        self.assertIn(b"Certificate of Participation", member_page.data)
        downloaded = member_client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(downloaded.status_code, 200)
        self.assertIn(
            "filename=ZIPSTUDENT001.pdf",
            downloaded.headers["Content-Disposition"],
        )
        with pymupdf.open(stream=downloaded.data, filetype="pdf") as stamped_pdf:
            page = stamped_pdf[0]
            self.assertIn(certificate_id, page.get_text())
            self.assertIn("Certificate No.", page.get_text())
            images = page.get_images()
            self.assertGreaterEqual(len(images), 1)
            qr_rect = page.get_image_rects(images[0][0])[0]
            self.assertLess(qr_rect.x0, page.rect.width * 0.1)
            self.assertGreater(qr_rect.y1, page.rect.height - 110)
        self.assertIn(b"Certificate of Participation", member_page.data)
        verification = member_client.get(f"/certificate/{certificate_id}")
        self.assertEqual(verification.status_code, 200)
        self.assertIn(b"Certificate verified", verification.data)
        self.assertIn(b"Zip Student", verification.data)
        self.assertIn(certificate_id.encode(), verification.data)
        lookup = member_client.get(
            f"/verify-certificate?certificate_id={certificate_id}"
        )
        self.assertEqual(lookup.status_code, 200)
        self.assertIn(b"Certificate verified", lookup.data)
        invalid_lookup = member_client.get(
            "/verify-certificate?certificate_id=FTI26-NOT-A-REAL-CERTIFICATE"
        )
        self.assertEqual(invalid_lookup.status_code, 200)
        self.assertIn(b"Certificate not found", invalid_lookup.data)
        downloaded.close()

        leader_client = self.app.test_client()
        self.assertEqual(leader_client.post("/login", data={
            "username": "zip_certificate_leader",
            "password": "ZipLeaderPassword123!",
            "role": "team_leader",
        }).status_code, 302)
        leader_page = leader_client.get("/team")
        self.assertIn(b"Certificate of Participation", leader_page.data)
        leader_download = leader_client.get(
            f"/certificate/{leader_certificate_id}/download"
        )
        self.assertEqual(leader_download.status_code, 200)
        self.assertIn(
            "filename=ZIPLEADER001.pdf",
            leader_download.headers["Content-Disposition"],
        )
        with pymupdf.open(
            stream=leader_download.data,
            filetype="pdf",
        ) as leader_stamped_pdf:
            self.assertIn(leader_certificate_id, leader_stamped_pdf[0].get_text())
            self.assertIn(
                "Team leader certificate",
                leader_stamped_pdf[0].get_text(),
            )
        leader_download.close()

        other_client = self.app.test_client()
        self.assertEqual(other_client.post("/login", data={
            "username": "FT_02",
            "password": "ChangeMe123!",
            "role": "team_leader",
        }).status_code, 302)
        self.assertEqual(
            other_client.get(f"/certificate/{certificate_id}/download").status_code,
            403,
        )

    def test_pdf_certificate_import_can_select_recipients_from_student_list(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            selected = User(
                username="SELECTED001",
                role="team_member",
                display_name="Selected Student",
                team_id=team.id,
            )
            selected.set_password("SelectedStudentPassword123!")
            other = User(
                username="NOTSELECTED001",
                role="team_member",
                display_name="Other Student",
                team_id=team.id,
            )
            other.set_password("OtherStudentPassword123!")
            db.session.add_all([
                TeamMember(
                    team_id=team.id,
                    full_name="Selected Student",
                    prn="SELECTED-PRN",
                    user=selected,
                ),
                TeamMember(
                    team_id=team.id,
                    full_name="Other Student",
                    prn="NOTSELECTED-PRN",
                    user=other,
                ),
            ])
            db.session.commit()
            selected_user_id = selected.id
            certificate_type_id = CertificateType.query.filter_by(
                name="Certificate of Participation"
            ).one().id

        def make_pdf(text):
            document = pymupdf.open()
            page = document.new_page()
            page.insert_text((72, 72), text)
            return document.tobytes()

        archive_buffer = io.BytesIO()
        with zipfile.ZipFile(archive_buffer, "w") as archive:
            archive.writestr("SELECTED001.pdf", make_pdf("Selected student"))
            archive.writestr("NOTSELECTED001.pdf", make_pdf("Other student"))
        archive_buffer.seek(0)

        self.assertEqual(self.login("admin", "Admin@12345!").status_code, 302)
        admin_page = self.client.get("/admin/certificates")
        self.assertIn(b"Select recipients from active student list", admin_page.data)
        self.assertIn(b"Selected Student", admin_page.data)
        self.assertIn(b"SELECTED-PRN", admin_page.data)
        response = self.client.post(
            "/admin/certificates",
            data={
                "action": "upload_files",
                "type_id": str(certificate_type_id),
                "pdf_recipient_scope": "selected",
                "pdf_recipient_user_ids": str(selected_user_id),
                "certificate_files": (archive_buffer, "student_certificates.zip"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            certificates = Certificate.query.filter_by(
                type_id=certificate_type_id
            ).all()
            self.assertEqual(len(certificates), 1)
            self.assertEqual(certificates[0].user_id, selected_user_id)
            self.assertEqual(certificates[0].dispatch_status, "pending")
            certificate_id = certificates[0].certificate_id
        self.assertEqual(
            self.client.post(
                "/admin/certificates/dispatch",
                data={"scope": "user", "user_id": str(selected_user_id)},
            ).status_code,
            302,
        )
        student_client = self.app.test_client()
        self.assertEqual(student_client.post("/login", data={
            "username": "SELECTED001",
            "password": "SelectedStudentPassword123!",
            "role": "team_member",
        }).status_code, 302)
        response = student_client.get(f"/certificate/{certificate_id}/download")
        self.assertEqual(response.status_code, 200)
        with pymupdf.open(stream=response.data, filetype="pdf") as stamped_pdf:
            self.assertIn("Certificate No.", stamped_pdf[0].get_text())
            self.assertIn(certificate_id, stamped_pdf[0].get_text())
        response.close()

    def test_member_can_update_own_profile_and_public_announcements_render(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            account = User(username="profile_member", role="team_member", team_id=team.id)
            account.set_password("ProfilePassword123!")
            db.session.add(TeamMember(
                team_id=team.id,
                full_name="Profile Student",
                prn="PROFILE001",
                user=account,
            ))
            db.session.commit()
        self.assertEqual(self.client.post("/login", data={
            "username": "profile_member",
            "password": "ProfilePassword123!",
            "role": "team_member",
        }).status_code, 302)
        self.assertEqual(self.client.post("/member/profile", data={
            "full_name": "Updated Student",
            "email": "updated@example.edu",
            "department": "Engineering",
        }).status_code, 302)
        with self.app.app_context():
            self.assertEqual(TeamMember.query.filter_by(prn="PROFILE001").one().full_name, "Updated Student")
        self.assertEqual(self.client.get("/announcements").status_code, 200)

    def test_all_roles_can_edit_own_profile_without_changing_access(self):
        with self.app.app_context():
            team = Team.query.filter_by(team_code="FT_01").one()
            roles = ("team_leader", "team_member", "mentor", "coordinator", "judge")
            users = {}
            for role in roles:
                username = f"profile_{role}"
                user = User(
                    username=username,
                    role=role,
                    display_name=f"Before {role}",
                    team_id=team.id if role in {"team_leader", "team_member"} else None,
                )
                user.set_password("ProfilePassword123!")
                db.session.add(user)
                if role == "team_member":
                    db.session.add(TeamMember(
                        team_id=team.id,
                        full_name=user.display_name,
                        prn="PROFILE_MEMBER_1",
                        user=user,
                    ))
                users[role] = username
            db.session.commit()

        for role, username in users.items():
            client = self.app.test_client()
            self.assertEqual(client.post("/login", data={
                "username": username,
                "password": "ProfilePassword123!",
                "role": role,
            }).status_code, 302)
            self.assertEqual(client.get("/account/profile").status_code, 200)
            response = client.post("/account/profile", data={
                "display_name": f"Updated {role}",
                "email": f"{role}@example.edu",
                "phone": "+1 (555) 010-1234",
                "department": "Engineering",
                "year": "2026",
            })
            self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            for role, username in users.items():
                user = User.query.filter_by(username=username).one()
                self.assertEqual(user.display_name, f"Updated {role}")
                self.assertEqual(user.email, f"{role}@example.edu")
                self.assertEqual(user.phone, "+1 (555) 010-1234")
                self.assertEqual(user.role, role)
                if role in {"team_leader", "team_member"}:
                    self.assertEqual(user.team_id, Team.query.filter_by(team_code="FT_01").one().id)
                if role == "team_member":
                    self.assertEqual(user.member_record.full_name, "Updated team_member")
                    self.assertEqual(user.member_record.department, "Engineering")
                    self.assertEqual(user.member_record.email, user.email)

    def test_profile_rejects_invalid_contact_details(self):
        with self.app.app_context():
            user = User(username="invalid_profile", role="judge", display_name="Original Name")
            user.set_password("ProfilePassword123!")
            db.session.add(user)
            db.session.commit()
        self.assertEqual(self.login("invalid_profile", "ProfilePassword123!").status_code, 302)
        response = self.client.post("/account/profile", data={
            "display_name": "Changed Name",
            "email": "not-an-email",
            "phone": "555-0100",
        })
        self.assertEqual(response.status_code, 400)
        with self.app.app_context():
            user = User.query.filter_by(username="invalid_profile").one()
            self.assertEqual(user.display_name, "Original Name")
            self.assertEqual(user.email, "")


if __name__ == "__main__":
    unittest.main()
