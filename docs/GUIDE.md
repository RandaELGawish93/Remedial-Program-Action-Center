# Remedial Program Action Center: illustrated guide

A walkthrough of the master workbook, the teacher and principal workbooks, every menu item and every email preview.

> Every picture in this guide is an **illustration filled with fictional sample data** (names like *Sarah Collins*, emails at `example.edu`). No real school, staff or student data appears anywhere in this repository.

## Contents

- [Menus](#menus)
- [Master workbook tabs](#master-workbook-tabs)
- [Teacher workbook](#teacher-workbook)
- [Principal workbook](#principal-workbook)
- [Email previews](#email-previews)
- [Messages](#messages)

## Menus

The master workbook has seven numbered menus, roughly in the order you use them during a term.

### Menus 1–4

![Menus 1–4](screenshots/remedial-menu-1.png)

**1. Setup** builds each tab of the master workbook: Settings, Student Roster, Scores Progress, the Dashboard shell, Email Log, Teacher Roster, Remedial Attendance (Teachers), MoM Log and HoD Directory. Each item can be re-run safely.

**2. Teachers & Principals**

| Item | What it does |
|---|---|
| Create Teacher Workbooks | Creates one workbook per remedial teacher (Profile, Progress Report Entries, Week 1, Week 5, Week 10) inside that teacher's folder. |
| Fill Teacher Folder Links | Writes each teacher's Drive folder link into the Teacher Roster. |
| Refresh Teacher Roster | Rebuilds the Teacher Roster from the Student Roster after changes. |
| Sync Teacher Profiles + Refresh Scores Progress | Pushes current-week info to every teacher Profile and pulls scores back into Scores Progress. |
| Send: Workbook + Guide (all / selected teacher) | Emails each teacher their workbook link and the guide, after a preview. |
| Send: Principal Workbooks | Emails each division principal their read-only workbook link, after a preview. |

**3. Audit & HoD Communication**: run the audit for a tab (Week 1, Week 5 + Progress Report Entries, Week 10), then send the matching *missing entries* notice to Heads of Department with a deadline. Also: send HoDs a status overview of their department's workbooks, check generated action plans against the template, and send thank-you emails for flawless tabs.

**4. MoM Log**: three meeting rounds per term. For each round, *Generate* (after marking attendance) creates one Minutes of Meeting document per teacher, and *Send* emails them.

### Menus 5–7

![Menus 5–7](screenshots/remedial-menu-2.png)

**5. Send Deadlines**: one email to every teacher with HoDs and principals in CC.

**6. Student Roster & Probation**: upload the signed probation letter for the selected student, and lock the roster so only the owner can edit it.

**7. System & End of Term**: install the periodic sync trigger (run once), delete all triggers if something goes wrong, generate the Division Reports (Week 11 only) and archive the trimester.

## Master workbook tabs

### Settings

![Settings](screenshots/remedial-master-settings.png)

Contacts (Academic Dean, principals, office assistants, directors…) with an Active switch, then the system configuration: term number, term start date, number of weeks, and the Drive folder and Google Docs template IDs used for action plans, progress reports, MoMs and archives.

### Student Roster

![Student Roster](screenshots/remedial-student-roster.png)

Every student in the program: grade, class, subject, program (Intervention or Prevention), tier, advisor, remedial teacher, starting score, status, the threshold score and links to the probation letter, action plan and progress report.

### Teacher Roster

![Teacher Roster](screenshots/remedial-master-teacher-roster.png)

Built from the Student Roster: each remedial teacher's subject, classes, number of students, folder links and workbook link.

### Scores Progress

![Scores Progress](screenshots/remedial-master-scores-progress.png)

Starting score and cumulative scores pulled from every teacher workbook, plus the final score entered by the program lead. Filters at the top narrow by subject, division or teacher.

### Dashboard

![Dashboard](screenshots/remedial-dashboard.png)

Colour-coded cards: students by tier and subject, teachers per division and subject, documents still missing (action plans and progress reports) and teachers who are off track.

### Remedial Attendance (Teachers)

![Remedial Attendance (Teachers)](screenshots/remedial-master-remedial-attendance-teachers.png)

Teacher punctuality for each remedial session from Week 2 to Week 10 (On Time, Late, Absent), grouped by division and subject.

### HoD Directory

![HoD Directory](screenshots/remedial-master-hod-directory.png)

Which Head of Department covers each subject and division, and whether they should receive emails.

### Audit — Week 1

![Audit — Week 1](screenshots/remedial-master-audit-week-1.png)

Created by an audit run. One row per teacher and student with the issue found (missing or placeholder entries, not started, generated with missing entries…), the exact fields that need attention and a *Send to HoD?* box.

### HoD Status Report

![HoD Status Report](screenshots/remedial-master-hod-status-report.png)

The same findings grouped by Head of Department, with links to each teacher's workbook and the send status.

### Generic Phrases

![Generic Phrases](screenshots/remedial-master-generic-phrases.png)

Your blocklist of vague phrases (*more practice*, *extra support*, *try harder*…). Entries that contain them are flagged.

### Generic Entry Review

![Generic Entry Review](screenshots/remedial-master-generic-entry-review.png)

Flagged entries waiting for your decision: the teacher, the student, the field, the text and a suggested note. Tick **Approve** to include it in the next email or **Ignore** to dismiss it. Nothing is emailed without your approval.

### MoM Log

![MoM Log](screenshots/remedial-master-mom-log.png)

Attendance for each meeting round, a link to each generated MoM PDF and whether the email was sent.

### Email Log

![Email Log](screenshots/remedial-master-email-log.png)

Every email sent: date, recipients, subject, type, related teacher or student, status and any error.

## Teacher workbook

Each teacher gets their own spreadsheet. Lead-only areas are protected; teachers fill in only their own fields.

### Profile

![Profile](screenshots/remedial-teacher-workbook-profile.png)

Read-only summary kept in sync by the master: current week, MoM status, this week's attendance, and a small dashboard of how many documents are still to be generated with their deadlines.

### Week 1 (Action Plan 1)

![Week 1 (Action Plan 1)](screenshots/remedial-teacher-workbook-week-1.png)

For each student the teacher writes the spotted weaknesses, steps to tackle them, means of assessment and evidence of improvement, then ticks **Generate?**. Within about 15 seconds the action plan PDF is created and linked, and the Status column confirms it (or lists the mandatory fields still empty).

### Week 5 (Action Plan 1 reflection and Action Plan 2)

![Week 5 (Action Plan 1 reflection and Action Plan 2)](screenshots/remedial-teacher-workbook-week-5.png)

Weekly exit-ticket scores and attendance for Weeks 1–4, the reflection on Action Plan 1 and the new plan. Only the first 12 columns are shown here.

### Week 10 (Action Plan 2 closeout)

![Week 10 (Action Plan 2 closeout)](screenshots/remedial-teacher-workbook-week-10.png)

Scores and attendance for Weeks 6–9 and the final reflection. Only the first 12 columns are shown here.

### Progress Report Entries

![Progress Report Entries](screenshots/remedial-teacher-workbook-progress-report-entries.png)

The teacher's comments for each student's progress report (arrives on time, completes work, needs support…). The master assembles them into one report per student.

## Principal workbook

### Principal dashboard

![Principal dashboard](screenshots/remedial-principal-workbook-dashboard.png)

One read-only workbook per division. The dashboard shows this week's remedial schedule with attendance, and a missing-work summary listing students whose action plan or progress report is still outstanding. The subject tabs are created empty for the principal's own notes.

## Email previews

Every send opens a preview first. The left column lists each email (untick to skip one); the right side shows exactly what will be sent. **Send this one to me (test)** sends a copy to you only.

### Workbook + guide to teachers

![Workbook + guide to teachers](screenshots/remedial-dialog-workbook-guide-all-teachers-preview-7.png)

Each teacher's personal workbook link and the guide.

### Principal workbooks

![Principal workbooks](screenshots/remedial-dialog-principal-workbooks-preview-3.png)

Each principal's read-only workbook link.

### Department workbooks status to HoDs

![Department workbooks status to HoDs](screenshots/remedial-dialog-department-workbooks-status-preview-3.png)

A status table for each department's teachers, with details of what is missing.

### Week 1 missing entries to HoDs

![Week 1 missing entries to HoDs](screenshots/remedial-dialog-week-1-action-plan-1-hod-notices-preview-3.png)

Sent after the Week 1 audit, with the deadline and exactly what each teacher still needs to fill in.

### Deadline to teachers

![Deadline to teachers](screenshots/remedial-dialog-deadline-week-1-action-plan-1-preview-7.png)

A deadline reminder to each teacher, with their own outstanding items and HoDs and principals in CC.

### Generate MoM

![Generate MoM](screenshots/remedial-dialog-generate-mom-round-2.png)

Before generating a round of MoMs: a general note for every document, each teacher's attendance and student summaries, and a box for your own observations per teacher.

### Send MoM

![Send MoM](screenshots/remedial-dialog-send-mom-round-1.png)

Choose the CC list and add an optional note before the MoM PDFs are emailed.

### Upload probation letter

![Upload probation letter](screenshots/remedial-dialog-upload-probation-letter.png)

Upload the signed letter for the selected student. It is saved to the probation folder and linked on the roster.

## Messages

### Audit summary

![Audit summary](screenshots/remedial-alert-audit-summary.png)

After an audit: how many teachers have not started, have missing entries or are complete.

### Workbooks created

![Workbooks created](screenshots/remedial-alert-workbooks-created.png)

After **Create Teacher Workbooks**.

### Periodic sync installed

![Periodic sync installed](screenshots/remedial-alert-periodic-sync.png)

The single time-driven trigger that processes every *Generate?* box.

### Refresh principal workbooks

![Refresh principal workbooks](screenshots/remedial-alert-principal-refresh.png)

Asked before principal workbooks are sent.

### HoD present?

![HoD present?](screenshots/remedial-alert-mom-hod-present.png)

Asked while generating a MoM, so the document records who attended.

### Division reports

![Division reports](screenshots/remedial-alert-division-reports.png)

Division reports are only available in Week 11, at the end of the term.

### Archive confirmation

![Archive confirmation](screenshots/remedial-alert-archive-confirm.png)

Explains what archiving does before anything is cleared.

---

[← Back to the README](../README.md)
