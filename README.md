# FUTURE-TECH INNOVISION 2026

Flask + SQLAlchemy exhibition management application starter with:
- One common login with a role selector and role-based redirects
- CSRF protection for state-changing forms and IP-based failed-login throttling
- Password change, admin password reset and account activation controls, with show/hide password buttons
- Initial team accounts FT_01 through FT_27
- Five technical domains
- Admin team/project management
- Admin-created team leader accounts when adding teams; deletion is blocked when score or attendance history must be preserved
- Admin-created additional administrator accounts
- Admin-managed judge-to-team assignments enforced by the backend
- Admin-managed mentor/team and coordinator/team assignments with one or two teams per account
- Admin-created team-leader and student accounts linked to a selected FT team
- Team-member registration with unique PRN checks and separate member accounts
- Imported student roster for FT_01–FT_27; each student's username is their Student_ID and initial password is their Team_ID
- Team leaders can submit an MP4 video, PPT and project documentation; assigned teammates, mentors, coordinators and authenticated admins can access the team files
- Admin team monitor provides event-wide access to all team/member/project details and submitted files; admin account management does not impersonate users or bypass authentication
- Admin-uploaded college/club logos and configurable college header
- Judge evaluation (60 marks)
- Admin evaluation (40 marks)
- Mentor-proposed 40-mark evaluations with admin review, score adjustment and final approval; coordinators can view scores
- Club-wide and team-project tasks with role/account targeting, reference files, assignee document/code submissions, per-assignee status and private file downloads
- Admins can create club or project tasks; team leaders and assigned mentors/coordinators can assign work within their projects
- Task assignments create detailed in-app notifications, project group-chat announcements, and optional SMTP email and MSG91 SMS notifications
- Open tasks due today or tomorrow receive once-per-day in-app and group-chat reminders, with optional email reminders
- Team-scoped group chat and private direct messages for leaders, members, assigned mentors and coordinators
- WhatsApp-style chat inbox with one-second message refresh, private administrator-to-account conversations searchable by name/username/email and filterable by role (including other admins), adjustable message text size, private file attachments of any type up to 25 MB, and sender-only five-minute message editing
- Staff dashboards show only approved evaluation marks; team leaders and members see scores after final leaderboard publication
- Domain-wise ranking page with five tables
- Top-three-per-domain nomination generation that supports domains with fewer than three eligible teams, including empty domains
- Per-domain result announcements, rank boards and notifications to team leaders, members, mentors and coordinators
- Project video submission during registration or project updates
- Final winner selection (one per domain)
- Announcements and basic audit logging

## 1. Install
Python 3.10+ recommended.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
```

## 2. Configure
Copy `.env.example` to `.env`. `run.py` loads the file automatically. For local first run, SQLite works without setup.
For MySQL, create a database named `future_tech_innovision`, set DATABASE_URL in `.env`, e.g.
`mysql+pymysql://root:YOUR_PASSWORD@localhost/future_tech_innovision`

Load environment variables before running (PowerShell):
```powershell
$env:FLASK_APP="run.py"
```
Or install python-dotenv integration in your shell; simplest local run:
```bash
python run.py
```

## 3. Open
Visit http://127.0.0.1:5000

## 4. Initial login credentials
- Admin: use the username and password configured with `BOOTSTRAP_ADMIN_USERNAME` and `BOOTSTRAP_ADMIN_PASSWORD` in your private `.env` file. The default local admin username is `24062231242187`.
- Team accounts: usernames `FT_01` through `FT_27`. If `TEAM_INITIAL_PASSWORD` is unset, each receives a unique random password; an Admin can issue/reset credentials in User Account Management.

Production startup requires a private `BOOTSTRAP_ADMIN_PASSWORD` and secure cookie configuration. When migrating an existing database, the legacy `admin` account is renamed to the configured admin username and its password is set from `BOOTSTRAP_ADMIN_PASSWORD`; if both accounts exist, the legacy account is deactivated.

## 5. Initial setup checklist
1. Login as admin.
2. Open Teams & Domains.
3. Assign each team to one of the five domains.
4. Enter team leader, project and registration details.
5. Reset each team account password and create mentor, coordinator and judge accounts from User Account Management.
6. When creating team leaders or members, select their FT team; member accounts also require a unique PRN.
7. Assign one or two FT teams to each mentor and coordinator from Mentor and coordinator assignments.
   Existing coordinators assigned by domain must be reassigned to teams after upgrading.
8. Team leaders can also register student members from their own dashboard; each member receives a separate login.
9. Configure college and club branding from the College and club branding page.
10. Create and configure judge assignments before live judging.
11. Review the 60-mark and 40-mark criteria.
12. Leaders can upload an MP4 with their project details; assigned team members, mentors and coordinators can view the submission.
13. Generate up to three nominations per domain. Empty or smaller domains do not block nomination publication. Publish nominations, then announce each domain's rank board separately from the admin dashboard.
14. Select one winner in each domain. The admin can then publish the final all-domain winner announcement.
15. Test with sample teams before the actual exhibition.

### Task notifications

Task assignments and submissions appear in the signed-in user's Tasks page. Assignment notices include the task details, assigner, deadline, and workspace link; project task assignments are also announced in the team's group chat. Admins can assign club-wide tasks to selected roles, selected accounts, or all active accounts. Admins can create project tasks from Tasks; team leaders and assigned mentors/coordinators can create project tasks in their team workspace. Tasks accept reference files, and assignees can submit documentation or code files (up to 25 MB per file).

Open tasks due today or tomorrow generate an in-app notice and, for project tasks, a group-chat reminder. Email reminders are sent when SMTP is configured. Reminders are deduplicated per task and calendar day and are processed on incoming application requests; configure a scheduled request or external scheduler if reminders must run at a guaranteed time during periods with no site traffic.

Task email and SMS are not sent until their provider settings are configured in the private deployment environment; the default contact email and phone are not delivery credentials. For email, configure `TASK_EMAIL_SMTP_HOST`, `TASK_EMAIL_SMTP_PORT` (587 for STARTTLS or 465 for implicit TLS), `TASK_EMAIL_SMTP_USERNAME`, `TASK_EMAIL_SMTP_PASSWORD`, and optionally `TASK_EMAIL_FROM`. If the SMTP server uses unauthenticated relay, leave both username and password empty. Restart the application after changing environment settings. Admins can check provider readiness and contact coverage on Tasks, and update user email/phone in Accounts. Team-member registration also copies the entered contacts to the member login account.

Task assignment SMS uses MSG91's [Send SMS flow API](https://docs.msg91.com/sms/send-sms.md). Set `TASK_SMS_AUTH_KEY` and `TASK_SMS_TEMPLATE_ID` privately. The template must be approved and configured in MSG91/DLT; by default the application substitutes the task title into the template variable `VAR1`. Set `TASK_SMS_VARIABLE_NAME` to match a differently named template variable exactly, and `TASK_SMS_COUNTRY_CODE` to change the default country prefix `91` for local 10-digit profile numbers. The API request is sent only when both the auth key and template ID are configured. Users need valid phone numbers in their account profiles; SMS errors are logged and surfaced without rolling back task creation. A provider acceptance response means MSG91 accepted the request, not that the carrier delivered it.

The default club contact is `future-techclub@orchidengg.ac.in` and `9834599110`; these are contact details, not recipient addresses.

## Student roster logins
The supplied roster is stored in `data/student_roster.csv`. Validate it with `python scripts/import_student_roster.py`, then import it into the configured database with `python scripts/import_student_roster.py --apply`. The import is safe to rerun: existing accounts keep their current passwords, and the original seeded FT team-leader logins are converted to their listed leader Student_ID accounts.

Students sign in with their `Student_ID` as the username and their `Team_ID` (for example `FT_01`) as the initial password. New and reset passwords must be at least 4 characters; there is no maximum length.

Judges submit their marks for admin approval; admins can approve, adjust, return them for correction, or reconfirm an existing approval. The **Judge evaluations & approval** dashboard view lists each submitted judge evaluation with its marks, approval status, and admin actions; **Admin 40-mark entries** lists only admin-entered scores out of 40; the **100 marks Evaluation** remains the team-level score summary. Mentor 40-mark proposals and entered judge scores appear in the summary before approval, while only approved judge scores and the finalized admin 40-mark score count toward rankings. Dashboard count cards open detail lists for the corresponding mentors, coordinators, teams, members, videos, winners, and evaluations.

Admins can open **Data & backups** to download a ZIP containing a SQLite database snapshot when SQLite is configured, CSV exports for application tables, the database schema, and uploaded files. Password hashes are redacted from the user CSV, but remain in the database snapshot; protect the downloaded archive as sensitive data. Technical-domain cards open a domain roster with team projects, leaders, and members. Detailed contact, staff assignment, and evaluation data is available only to admins.

Nominations rank up to three eligible teams per domain using approved judge and admin scores. A domain with one or two eligible teams nominates those teams; a domain with no eligible teams remains empty and does not prevent generation or publication. Published nominations remain frozen. The admin can use **Cancel published nominations and results** to withdraw nomination/domain/final publication, remove selected winners and related in-app result notifications, preserve scores, and return nominations to draft so scores can be corrected and results regenerated. Notifications already seen outside the app cannot be recalled. Public rankings and role dashboards disclose a domain only after the admin announces its results.

In User Account Management, reusing a username updates its role only when both the selected role and password differ from the existing account. Reusing the same role or password does not update the account; use the password-reset action to change only a password.

## 6. Production deployment (Render)

The included `render.yaml` defines a Docker-based production web service, managed PostgreSQL database, persistent upload disk, and an ONLYOFFICE Docs image service for in-page Word editing. The app Docker image installs LibreOffice for rendering Word certificate templates to PDF. ONLYOFFICE runs as a separate Render service with a 2 CPU / 4 GB plan; it is an additional billable service. Production requires PostgreSQL; SQLite remains for local development and tests. The app normalizes Render's PostgreSQL connection URL, checks database connectivity at `/healthz`, and stores uploaded project files, videos, chat/task attachments, branding, and certificate artwork under `UPLOAD_FOLDER`.

1. Create a **private** GitHub repository and publish the project source from VS Code. Keep `.env`, the local database under `instance/`, `uploads/`, and the student roster CSV out of Git; `.gitignore` excludes these. Do not put student/contact spreadsheets or credentials in a repository.
2. Create a Render account, choose **New → Blueprint**, connect the private repository, and let Render read `render.yaml`. The service is configured to run one Gunicorn worker because its persistent upload disk is attached to one instance.
3. When prompted for `BOOTSTRAP_ADMIN_USERNAME` and `BOOTSTRAP_ADMIN_PASSWORD`, set your administrator login. Use a unique password with at least 12 characters. `SECRET_KEY` is generated by Render; keep it private and do not rotate it casually, since rotation signs users out.
4. Choose a Render plan that supports both the persistent disk and managed PostgreSQL. The blueprint also provisions a separate 2 CPU / 4 GB ONLYOFFICE document server; review its current cost before applying the Blueprint. The app's ONLYOFFICE URL and JWT secret are wired from that service automatically. The blueprint asks for a 20 GB upload disk. Do not remove the disk or change its mount path after users upload data. Keep the app linked to the same PostgreSQL database during redeploys.
5. Wait for the app and ONLYOFFICE services to report healthy, then open the app's HTTPS URL. In Certificate studio, open a saved Word template, choose **Edit Word certificate here**, edit and save in ONLYOFFICE, then close the editor to sync the Word file and PDF preview. Each saved Word edit returns the template to draft; approve it again before issuing certificates.
6. Sign in as admin, reset the generated FT team-account passwords, create staff accounts through User Account Management, and have team leaders register student members or securely arrange a roster import. Test one account per role before sharing links.
7. Share the Render HTTPS URL and each person's own username/password privately. Students should sign in using their individual accounts to edit their permitted profile details and submit projects; the app associates submissions and notifications with the signed-in user/team, so other accounts should not see or edit private records.
8. Before launch, submit a test project and attachment, sign in as the intended reviewer, verify visibility, redeploy once, and verify the data and files still exist. A successful `/healthz` response checks the app-to-database connection, not permission correctness or backup integrity.
9. Enable the database provider's automated backups and periodically test a restore. Also download **Data & backups** from the admin area and store it privately off the server; it includes uploaded files and table exports, but its CSV export redacts password hashes and is not a replacement for a native PostgreSQL backup. Protect backups as sensitive personal data.

### Task assignment email

Task assignments already create in-app notifications and post project tasks to the project group chat. To email assignees too, configure SMTP for a mailbox authorized to send as `future-techclub@orchidengg.ac.in`. The Render blueprint sets this as `TASK_EMAIL_FROM` and prompts for `TASK_EMAIL_SMTP_HOST`, `TASK_EMAIL_SMTP_USERNAME`, and `TASK_EMAIL_SMTP_PASSWORD`; enter the mail provider's SMTP host and authorized mailbox credentials or app password in Render's private environment settings. The default SMTP port is 587 with STARTTLS; use 465 only if your provider requires implicit TLS. A `noreply` alias is usable only if the mail provider authorizes it as a sender for that SMTP account. Never put mailbox passwords or app passwords in Git, source code, or chat. Email is sent only to assignees with an email address in their account/profile or student roster. Until SMTP is configured, tasks still appear in app notifications and project group chat.

Admins can publish an announcement with **Send an in-app notification to every active account** enabled. Each active account gets a private notification, an unread badge, and an in-app message banner on their next page view; users can mark it read or review it later in Account History. This uses the app database and does not depend on SMTP. It is not a mobile push notification or an email.

The hosted database and persistent disk are separate from this development folder. Existing local records in `instance/future_tech_innovision.db` and files in `uploads/` do **not** transfer automatically when you first deploy. Do not delete the local copies; a planned database-and-files migration is needed if you want to carry existing entries into the hosted system. New hosted entries persist across normal application redeploys as long as you retain the same database and disk.

To copy the local SQLite application records into an already-deployed PostgreSQL database, run `.\.venv\Scripts\python.exe scripts\migrate_local_sqlite_to_postgres.py` from PowerShell on Windows (or `python scripts/migrate_local_sqlite_to_postgres.py` on macOS/Linux). The script asks for the Render admin username/password and Neon PostgreSQL connection string without echoing passwords or the connection string. It verifies the local admin account, then requires typing `REPLACE` before it replaces application tables in the destination database in one transaction. This is destructive to data currently in the hosted application; make a backup first and run it only against the intended database. The script copies database records only—not files from `uploads/`. Re-upload branding and certificate assets separately. Render Free has no persistent disk, so uploaded files there may be lost after a restart or redeploy.

This workspace has not been published or deployed: deployment needs your private GitHub repository and Render account. Never paste passwords or database URLs into chat; enter them only in Render's private environment settings.

## 7. Tests
Run core workflow and authorization tests with:

```bash
python -m unittest discover -s tests -v
```

## 8. Implementation status / next stages
This is a connected functional application foundation, not a fully production-certified deployment. Implemented workflows include member registration and self-profile editing, role-scoped team/project/video views, coordinator and judge attendance, mentor score proposals, admin-approved final scoring, team-scoped task assignment and group/direct chat, searchable private administrator chats, one-second message refresh, private chat attachments and five-minute sender-only message editing, final-leaderboard score visibility for teams, per-domain result announcements and rank boards, and team/admin account management. Login-attempt throttling, CSRF protection, minimum-length password controls, judge assignment enforcement, evaluation locking, and nomination/publication guards are also implemented.

Before the live event:
- Add email-based password recovery and decide whether rate limits should be stored in a production shared cache.
- Add an administrator review/approval workflow for student profile changes.
- Configure video size/duration limits and private object storage if deployed online.
- Add fixed male voice audio/TTS assets and the full timed presentation player.
- Add detailed audit coverage, backups, HTTPS and production deployment settings.
- Configure and publish the tie-break rule before judging; current multi-judge scores use the average of all finalized assigned judges.
- Test all workflows and permissions using test accounts before real use.

The system deliberately leaves college/club logos, event dates, actual team names and domain allocation configurable.

### Certificate studio

Admins can open **Certificate studio** from the admin dashboard to create and approve templates for participation, winner, mentor, coordinator, and club-head/judge certificates. Upload a PDF or image as the design background, then position and style editable text fields, cover areas, certificate numbers, QR verification codes, and custom text. Cover areas are drawn over the selected part of the background before the recipient details are added; old text in a PDF or image is not detected or removed automatically.

Admins can also generate and immediately dispatch a demo participation or winner certificate to one selected active student account. It appears in that student's dashboard and in the admin issued-certificate list; email is sent when configured. The admin can delete that certificate from the list. Admins can upload a Word (`.docx`) source, preview its extracted text and rendered PDF, and use it without a separate PDF/image background. The deployed ONLYOFFICE Docs service lets admins edit Word text, images, fonts, and layout directly in the certificate studio; saved changes update the DOCX and its rendered PDF preview and return the template to draft for re-approval. In the studio, admins can select existing Word text and map it to a recipient field, or use placeholders. Supported aliases include `<Name>`, `<Role>`, `<Domain>`, `<Event>`, `<Day>`, `<Date>`, `<Venue>`, `<Team>`, `<Title>`, and the `{{recipient_name}}`-style field names. Role and event day can be mapped from spreadsheet columns. Administrators can select each PDF merge field to set its font, size, color, bold, and italic styling, and set the same text styles per Word placeholder. The QR marker can be dragged to set its position in both output formats; each certificate number is printed directly below its QR code. Each personalized Word download gets its own certificate number and unique verification QR. After importing recipient rows from Excel (`.xlsx`) or CSV, each generated certificate has an individualized PDF and, when a Word source is present, a personalized DOCX. Word-based PDF rendering requires LibreOffice (included in the Render Docker image), and each Word certificate must render to exactly one page. Generated certificates remain pending and hidden from recipients until an admin approves and dispatches them to all accounts, a domain, or one account. Re-importing recipients updates pending drafts; already-dispatched certificates remain available. Admins can delete an individual certificate or all issued records. Text already printed in a Word source can be selected and mapped or replaced with placeholders. Optional email delivery occurs at dispatch and requires SMTP configuration. Existing certificates remain dispatched during the schema upgrade, and certificate data and files are preserved across application restarts.

Admins may also import pre-designed student certificate PDFs from **Import finished student certificate PDFs**. Select multiple PDF files or a ZIP; each filename must be the student's exact `Student_ID.pdf`/PRN, such as `STUDENT123.pdf`. Like the Mailings recipient selector, the importer can use all account IDs found in the filenames or a selected list of active student accounts; selected recipients are matched by Student ID/username or PRN, and only their matching PDFs are imported. Each imported PDF receives a unique verification QR and certificate number below it in a compact block at the lower-left of its first page, and becomes a pending certificate. After dispatch, scanning the QR opens a public verification record showing the recipient, certificate type, team, issuer, and issue date. The public **Verify Certificate** page also lets anyone look up a certificate number manually; deleted or unknown numbers are reported as not found and pending certificates are not reported as issued. Dispatch them to the intended student, domain, or all pending recipients; dashboard delivery is available without email, and configured SMTP attaches the stamped PDF.

Verification QR links use the deployed Render service URL automatically when available. For another host or a custom domain, set `CERTIFICATE_PUBLIC_BASE_URL` to the canonical HTTPS origin before creating certificates. Create final certificates after deploying, or configure this base URL first, so QR links lead outside the local development server.
