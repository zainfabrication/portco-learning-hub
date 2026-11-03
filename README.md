# PortCo Learning Hub – Dashboard

A live dashboard that tracks every portco, content project, communication and file.
**All numbers come from `data.xlsx`.** Update the Excel file and the dashboard updates itself.

## What's in this folder

| File | Purpose |
|---|---|
| `index.html` | The dashboard page |
| `assets/style.css` | Look and feel |
| `assets/app.js` | Reads the Excel file and builds the charts and tables |
| `data.xlsx` | Your data – the only file you need to edit |

## Set it up on GitHub (one time, about 5 minutes)

1. Sign in to GitHub and click **New repository**. Name it e.g. `portco-learning-hub`.
   > Make it **Private** if the data is confidential (GitHub Pages on private repos needs a paid/enterprise plan – check with IT).
2. Click **Add file → Upload files** and upload everything in this folder (keep the `assets` folder).
3. Go to **Settings → Pages**. Under *Build and deployment* choose **Deploy from a branch**, branch **main**, folder **/(root)**, then **Save**.
4. After about 1 minute your dashboard is live at
   `https://<your-github-username>.github.io/portco-learning-hub/`

## How updating works

1. Open `data.xlsx`, update the numbers or add rows, and save.
2. In GitHub, open the repo → **Add file → Upload files** → upload the new `data.xlsx` → **Commit changes**.
   (Or use GitHub Desktop: save the file in the repo folder → Commit → Push.)
3. The dashboard checks for a new version **every 60 seconds** and refreshes on its own.
   GitHub Pages usually publishes the change within 1–2 minutes. Click **Refresh** to check right away.

**Quick preview without GitHub:** open the dashboard and click **Load Excel** to pick any `.xlsx` on your computer.
It shows that file instantly (nothing is uploaded anywhere).

## Excel structure (keep the headers in row 1 exactly as they are)

| Sheet | Columns |
|---|---|
| **Portcos** | PortcoID, PortcoName, Industry, SPOC, SPOCEmail |
| **Projects** | ProjectID, PortcoName, ProjectName, ContentType, IDOwner, ScriptReceived, StartDate, DueDate, Status, Progress, FeedbackRounds, Priority |
| **Communications** | Date, PortcoName, ProjectID, Type, Direction, Subject, Owner |
| **Files** | FileName, PortcoName, ProjectID, FileType, Version, UploadedOn, Link |

- **Status** values: Script Received, In Development, In Review, Portco Feedback, Delivered, On Hold
- **Direction**: Sent or Received
- **Progress**: 0–100
- **PortcoName** must match exactly across sheets.
- **Link** (optional): paste a SharePoint link to make the file clickable.

## Health flags (calculated automatically)

| Flag | Rule |
|---|---|
| 🔴 Overdue | Due date has passed and status is not Delivered |
| 🟠 At risk | Due within 7 days and progress below 70% |
| 🟢 On track | Everything else that is active |
| 🔵 Delivered / ⚪ On hold | Taken from Status |

Change these rules in `assets/app.js` → `CONFIG` at the top of the file.

## Settings you can change (`assets/app.js` → `CONFIG`)

- `dataFile` – name or URL of the Excel file
- `autoRefreshSeconds` – how often to check for updates (default 60)
- `dueSoonDays` – the "at risk" window (default 7)

## Data privacy

The sample `data.xlsx` uses **fictional** portco names. Before you add real portco data, confirm with IT
that it's OK to host it on GitHub. A **public** repo makes the data visible to anyone.
For internal use, the same files can also be hosted on a SharePoint site or an internal web server.
