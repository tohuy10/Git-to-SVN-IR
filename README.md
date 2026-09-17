# Git-to-SVN-IR (GitHub Review Log Collector & SVN Sync)

An automated tool and Jenkins CI/CD pipeline that aggregates code review activities across multiple GitHub repositories, formats them into a structured Excel tracking report, and synchronizes the output directly to a designated Subversion (SVN) repository path.

---

## Current Workflow & Motivation

1. **Current State:**
   - Every week, Tech Leads, Devs, and team members who need to review past PR comments must navigate manually to every repository on GitHub, opening both the **Conversation** and **Files changed** tabs across dozens of Pull Requests to trace discussion history (Commit IDs, Dev names, Reviewer names, inline feedback, etc.).
2. **Key Challenges:**
   - **Time & Effort Consuming:** Manually tracking reviews across multiple active repositories and PRs drains considerable engineering hours weekly.
   - **Data Blindspots (Human Error):** High risk of missing critical comments buried deep in diff lines or review summaries on large PRs.
3. **Proposed Improvement:**
   - Automate data gathering via the GitHub REST API across all designated repositories.
   - Collect both **Main Review Summaries** (`/pulls/{id}/reviews`) and **Line-specific comments** (`/pulls/{id}/comments`).
   - Format, deduplicate, and continuously append records into a centralized Excel tracking document (`Git_Review_Log.xlsx`).
   - Automate execution via a scheduled weekly Jenkins job that commits the latest spreadsheet directly to the designated SVN repository path.

---

## Expected Outcomes

- **Effort Saving:** Eliminates manual checks of past Pull Requests across repos.
- **Complete Review History:** Captures all review records and line-by-line diff discussions without manual data omission risks.
- **Zero Manual Action:** Automated pipeline executes on schedule, maintaining up-to-date documentation on SVN automatically.

---

## Features

- **Multi-Repository Ingestion:** Aggregates logs across multiple repositories in a single run. Supports SSH (`git@github.com:...`), HTTPS URLs, and plain `owner/repo` slugs with automatic whitespace normalization.
- **Unified Chronological Timeline:** Pull Requests and internal discussions are ordered by creation timestamp ascending, intertwining main PR reviews and inline line comments seamlessly.
- **Dynamic Scan Timeframes:** Choose flexible collection windows on demand (`1 week`, `2 weeks`, `3 weeks`, `1 month`, `2 months`, `3 months`, `6 months`, `1 year`, `2 years`, `All History`, or `Incremental`).
- **Smart Incremental Sync (Watermarking):** Scans the existing Excel sheet on SVN for the latest recorded review date per repository. Fetches only PRs updated after that timestamp with a 10-minute safety buffer.
- **Persistent Outdated Line Citations:** Captures multi-line and single-line comment positions (e.g., `src/index.ts:L45` or `src/api.py:L10-L18`). Automatically falls back to `original_line` and `original_start_line`
so line references survive even after code is modified or threads are resolved.
- **HTTP Connection Pooling:** Uses a shared, pooled `requests.Session()` with HTTP Keep-Alive, significantly reducing TLS handshake overhead and network latency across hundreds of GitHub API requests.
- **GMT+7 Timezone Normalization:** All UTC timestamps from GitHub are automatically converted to GMT+7 and formatted uniformly as `DD Mon, YYYY HH:MM AM/PM` (e.g., `11 Sep, 2026 12:08 AM`).
- **Composite Deduplication:** Prevents duplicate rows when resyncing using a composite key:  
  `Repository | Commit ID | Reviewer Name | Created At | Comment Body`.
- **Self-Healing SVN Working Copy:** Verifies remote repository URLs via `svn info` before syncing. Automatically wipes and checks out a fresh copy if target paths or branches change, preventing SVN metadata collisions.
- **Jenkins CI/CD Native:** Includes a parameterized declarative `Jenkinsfile` with credentials masking.

---

## Output Structure (`Git_Review_Log.xlsx`)

The generated Excel workbook contains a styled sheet named **"Review Logs"**:

| Column | Header | Description |
|---|---|---|
| **A** | Repository | Repository slug (`owner/repo`) |
| **B** | Commit ID | Git commit SHA targeted by the review |
| **C** | Dev name | PR author / commit author name |
| **D** | Created at | Timestamp when review or comment was submitted (GMT+7) |
| **E** | Commit content | Git commit message |
| **F** | Reviewer name | GitHub username of the reviewer |
| **G** | Reviewer comment | The feedback, markdown body, or inline note |
| **H** | Review Type (Temporary column) | `PR Review` (general) or `Line Specific` (diff comment) |
| **I** | File Path / Line (Temporary column) | Target file and line location (e.g., `src/main.py:L12` or `—`) |
| **J** | Review State (Temporary column) | PR review state (`COMMENTED`, `APPROVED`, `CHANGES_REQUESTED`, or `—`) |

---

## Synchronization Modes & Timeframes

Scan windows are determined dynamically by the `SCAN_TIMEFRAME` parameter:

* **Relative Time Windows (Default: `1 week`):**  
  Computes a cutoff date based on current time (e.g., `1 week` $\rightarrow$ past 7 days, `1 month` $\rightarrow$ past 30 days, `1 year` $\rightarrow$ past 365 days). Inspects all PRs updated after the cutoff.
* **`Incremental (from latest auto-detected Excel Timestamp)`:**  
  Reads `Git_Review_Log.xlsx` directly from the SVN working directory, extracts the latest timestamp per repository, and fetches only newer PR updates.
* **`All History`:**  
  Traverses all pages of Pull Requests across repository history. Useful for initial repository onboarding or total report rebuilds.

---


## Prerequisites & Requirements

### Local Environment
- **Python 3.10+**
- **Subversion CLI client (`svn`)**
- Required Python libraries:
  ```bash
  pip install requests openpyxl
  # Or on Ubuntu 24.04+ (PEP 668 managed environments):
  sudo apt install -y python3-requests python3-openpyxl
  ```

---

## Jenkins Pipeline Setup

The pipeline is designed to work in locked-down Jenkins environments where developers do not have administrator permissions to add credentials under **Manage Jenkins**.

### 1. Job-Level Parameter Configuration

1. In Jenkins, open your Pipeline job and click **Configure**.
2. Under **General**, select **This project is parameterized**.
3. Add the following parameters:

| Parameter Type | Name | Example Value | Description |
|---|---|---|---|
| **String Parameter** | `REPO_LIST` | `https://github.com/org/repo1, https://github.com/org/repo2` | Comma-separated list of GitHub repositories (URLs or slugs). |
| **Password Parameter** | `GITHUB_TOKEN` | `ghp_xxxxxxxxxxxx` | GitHub Personal Access Token (`repo` read access). Concealed by Jenkins. |
| **String Parameter** | `SVN_URL` | `svn://172.16.3.43:3690/company_repo/Project_Engineering/6_Review/Git_review_log/` | Target SVN directory URL. |
| **String Parameter** | `SVN_USER` | `svnuser` | Username for SVN authentication. |
| **Password Parameter** | `SVN_PASS` | `password123` | Password for SVN authentication. Concealed by Jenkins. |
| **Choice Parameter** | `SCAN_TIMEFRAME` | `1 week`<br>`2 weeks`<br>`3 weeks`<br>`1 month`<br>`2 months`<br>`3 months`<br>`6 months`<br>`1 year`<br>`2 years`<br>`All History`<br>`Incremental (from latest auto-detected Excel Timestamp)` | Select review scan timeframe. `1 week` (first item) is the default for automated runs. |

4. Under **Build Triggers**, select **Build periodically** and configure the schedule:
   ```text
   H 2 * * 1
   ```
   *(Runs every Monday between 02:00 and 03:00 GMT+7)*.
5. Click **Save**.

> **Note on `parameters` in `Jenkinsfile`:**  
> The `parameters { ... }` block is intentionally omitted/commented out inside the `Jenkinsfile`. If declared in code with empty default values, Jenkins would overwrite your UI-saved concealed passwords on every Git push. Defining parameters via the Jenkins UI keeps passwords permanently encrypted on disk.

### 2. Manual vs. Automated Builds

- **Automated Scheduled Runs (Cron):** Jenkins runs in the background and automatically injects the concealed default passwords saved in the job configuration, defaulting to the top choice if there are multiple (`1 week`).
- **Manual Runs ("Build with Parameters"):** Jenkins injects saved defaults. Developers can select any scan window from the `SCAN_TIMEFRAME` dropdown before clicking **Build**.


---

## Project Structure

```text
.
├── Jenkinsfile        # Parameterized CI/CD pipeline definition
├── README.md          # Project documentation
└── sync_reviews.py    # Main synchronization and Excel formatting script
```
