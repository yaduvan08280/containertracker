
=======
# Container Tracker — Setup Guide
1
A Google Apps Script web app that tracks every container from booking to gate-in, with deadlines, Telegram alerts, admin approvals and analytics.

## Files

| File | Type in Apps Script editor | What it does |
|---|---|---|
| `Code.gs` | Script | Web entry, login, all API calls, setup |
| `Data.gs` | Script | Sheet layout + caching layer |
| `Notify.gs` | Script | Telegram messages + the 5-minute scheduler |
| `Shared.gs` | Script | Steps, fields, charge list, deadline maths (also sent to the browser) |
| `Index.html` | HTML | Page shell |
| `Styles.html` | HTML | Colours / layout (from your Dashboard_v1.18 palette) |
| `App.html` | HTML | The whole browser app |
| `appsscript.json` | Manifest | Timezone (Asia/Kolkata), permissions, web-app settings |

## Install (about 10 minutes)

1. Create a new **Google Sheet** (e.g. "Container Tracker DB"). Open **Extensions ▸ Apps Script**.
2. In the editor, delete the sample `Code.gs` content and create the files above. Use **+ ▸ Script** for `.gs` files and **+ ▸ HTML** for `.html` files. Name them exactly as shown, without the extension (`Code`, `Data`, `Notify`, `Shared`, `Index`, `Styles`, `App`).
3. Go to **Project Settings ⚙ ▸ Show "appsscript.json" manifest file in editor**, open `appsscript.json` and replace its contents with the one provided.
4. In the editor, select the function **`setup`** and click **Run**, then approve the permissions. This creates the 5 sheets, the 5 default users, the Drive upload folder and the 5-minute scheduler.
5. Go to **Deploy ▸ New deployment ▸ Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**. The app has its own login. If your Google Workspace blocks "Anyone", pick "Anyone within <your domain>".
6. Copy the web app URL and share it with the team. It works on phones too; they can "Add to Home Screen".

> **After any code change**, go to **Deploy ▸ Manage deployments ▸ ✎ ▸ Version: New version ▸ Deploy**. The URL stays the same.

## Default accounts

Every account must set its own password at first sign-in.

| User | Username | Temporary password | Role | Assigned to |
|---|---|---|---|---|
| Admin | `admin` | `Admin@123` | Admin | Everything |
| PC01 | `pc01` | `Pc01@123` | Process Coordinator | Steps 1–10 (reassign in Admin ▸ Workflow) |
| Harish Ji | `harish` | `Harish@123` | User | Shipment Tracking (ETD, BL date, ETA, Arrival, ISF, Filing, DO, Clearance, Arrival Notice) |
| Kuldeep Ji | `kuldeep` | `Kuldeep@123` | User | Charges & Expenses |
| Dhruv Ji | `dhruv` | `Dhruv@123` | User | Shipping Bill & BL (number + photo) |

## Telegram setup

1. In Telegram, message **@BotFather** ▸ `/newbot` and copy the **bot token**.
2. In the app: **Admin ▸ Telegram & alerts** ▸ paste the token ▸ **Save settings**.
3. Each user goes to **their name (bottom-left) ▸ Link Telegram ▸ Open Telegram ▸ Start ▸ Check**. They can also paste a chat id in Admin ▸ Users.
4. Add **mobile numbers** in Admin ▸ Users. Delay alerts tell the coordinator who to call and on which number.

**Who gets which message**

- **New task:** the users assigned to the next step, as soon as the previous step is done. Messages are grouped per booking.
- **Reminder:** the assignees, X hours before the deadline (default 2 h).
- **Delay alert:** Process Coordinators (and admins, if enabled). The alert includes the step, container, booking and each responsible person's name and mobile. It repeats every N hours (default 4) until the step is done. The responsible person also gets a nudge.
- **Approval needed:** admins, whenever someone asks to change an already-filled value.
- **Self reminders:** from the Reminders page. They can be one-off, daily or weekly.
- **Quiet hours** (default 22:00–07:00) hold back reminders and delay alerts. Self reminders still arrive.

## The pipeline and default deadlines

Admin can change every deadline in **Admin ▸ Workflow & deadlines**. A deadline is either "N hours after the previous step" or "a date field ± N hours".

| # | Step | Default deadline |
|---|---|---|
| 1 | Booking Details (Booking No., Line, Vessel, Party, Destination, No. of containers). Creates one pipeline per container. | — |
| 2 | Container Allotment Date | 24 h after booking |
| 3 | Container Allotment Details (Arrival date, Container No., Invoice No.) | the Allotment Date (end of day) |
| 4 | S.I & Gate-In Cut-offs + Receiving/Handover location | 24 h after step 3 |
| 5 | S.I Submission (mark done + remarks) | S.I Cut-off Date |
| 6 | Container Departure Date (Factory) | 24 h after step 5 |
| 7 | Stuffing / Dispatch (Est. reaching date, driver no., vehicle no., remark) | Departure Date |
| 8 | All Clearance Documents Preparation | 2 h after step 7 |
| 9 | All Clearance Documents Submission | 2 h after step 8 |
| 10 | Container Gate-In | Gate-In Cut-off Date |
| — | Shipping Bill & BL (mandatory after dispatch) | 24 h after step 7 |

For steps 2, 4 and 6, one tick on the form applies the same values to every container of that booking that is waiting at the same step.

## Rules built in

- **Edits need approval:** an empty field saves directly. Changing a value that is already filled creates a request in **Approvals**. Admin approves or rejects it, and the requester is told on Telegram. Admin's own edits apply immediately.
- **Only admin** can Close, Cancel or Reopen a container. Closing warns you if gate-in or SB/BL is missing.
- **History:** every action is written to the `Log` sheet and shown to everyone in the container's **History** tab.
- **Photos and documents:** photos are shrunk in the browser (max 1600 px) before upload, so they upload fast on mobile data. Files go to the Drive folder "Container Tracker — Uploads" and the **link** is written into the sheet column next to each charge. You can also paste an existing link (🔗).

## Sheets (5 only)

| Sheet | Contents |
|---|---|
| `Containers` | One wide row per container: booking info, all step fields, each step's done-at / done-by / deadline, tracking fields, SB/BL, 18 charges × (amount, photo 1, photo 2), 5 extra-expense slots, status |
| `Users` | Accounts, roles, mobile, Telegram chat id, assigned steps. Passwords are stored salted and hashed. |
| `Approvals` | Change requests and decisions |
| `Log` | Full history (also used for the History tab) |
| `Config` | Deadline rules, alert settings, self reminders |

Don't rename the header row or reorder columns by hand. Editing values by hand is fine: the app notices and reloads. **Admin ▸ Telegram & alerts ▸ Reload data from sheet** forces a reload.

## Why it's fast

- **Server side:** the whole dataset is cached in CacheService (chunked JSON). Reads never touch the sheet unless the cache is empty. Each write updates only the changed rows and refreshes the cache in the same locked operation.
- **Browser side:** the last snapshot is kept in the browser. The app draws immediately on open, then fetches **only the containers that changed**. It checks for changes every 45 s, which costs one tiny call when nothing changed.
- **Saves:** the screen updates first, and the server confirms in the background. If the server rejects a save, the screen rolls back.
- **Search, filters, tabs and analytics** all run in the browser on data it already has, so there is no waiting while typing.
- **Telegram:** messages are queued and sent after your save returns, so Telegram never slows down a save.

## Things to confirm with the team

- **Charge categories:** I grouped your 18 charges under the 5 headers in the order you listed them. Transportation (1), Shipping Line (9), Examination (3), After Gate-In (3), Destination (2). To move one, edit `chargeGroups` in `Shared.gs`.
- **Tracking fields:** ETD, BL Date, ETA, Actual Arrival and Filing Date are dates. ISF, DO, Clearance and Arrival Notice are free text, so a status or number can be written in them.
- **Step 7's estimated reaching date** is recorded and shown, but step 8's deadline follows your explicit "2 hours after step 7" rule. If you want it to use the estimated date instead, change it in Admin ▸ Workflow.

