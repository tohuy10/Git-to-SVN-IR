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

- **Multi-Repository Aggregation:** Synchronize reviews from multiple GitHub repositories in a single pipeline run.
- **Unified Chronological Timeline:** Pull Requests and internal discussions are ordered by creation timestamp ascending, intertwining main PR reviews and inline line comments seamlessly.
- **Accurate Line Citations:** Preserves exact code review file paths and line ranges (e.g., `README.md:L1-L2`).
- **Smart Incremental Sync (Watermarking):** Scans the existing Excel sheet on SVN for the latest recorded review date per repository. In normal operation, only PRs updated after that timestamp are fetched, cutting API traffic and execution time.
- **Full Historical Scan Mode:** Includes an in-code toggle (`SCAN_ALL_PRS`) to bypass watermarks and scan every historical Pull Request whenever a repository is onboarded or needs a full rebuild.
- **GMT+7 Timezone Normalization:** All UTC timestamps from GitHub are automatically converted to GMT+7 and formatted uniformly as `DD Mon, YYYY HH:MM AM/PM` (e.g., `11 Sep, 2026 12:08 AM`).
- **Composite Deduplication:** Prevents duplicate rows when resyncing using a composite key:  
  `Repository | Commit ID | Reviewer Name | Created At | Comment Body`.
- **Automated SVN Sync:** Automatically checks out, updates, styles, adds, and commits the resulting `.xlsx` document back to the target SVN path.
- **Jenkins CI/CD Native:** Includes a parameterized declarative `Jenkinsfile` with credentials masking.

---

## Output Structure (`Git_Review_Log.xlsx`)

The generated Excel workbook contains a styled sheet named **"Review Logs"**:

| Column | Header | Description |
|---|---|---|
| **A** | Repository | Repository slug (`owner/repo`) |
| **B** | Commit ID | Git commit SHA targeted by the review |
| **C** | Dev name | PR author / commit author name |
| **D** | Created at | Timestamp when review or comment was submitted (UTC) |
| **E** | Commit content | Git commit message |
| **F** | Reviewer name | GitHub username of the reviewer |
| **G** | Reviewer comment | The feedback, markdown body, or inline note |
| **H** | Review Type (Temporary column) | `PR Review` (general) or `Line Specific` (diff comment) |
| **I** | File Path / Line (Temporary column) | Target file and line location (e.g., `src/main.py:L12` or `—`) |
| **J** | Review State (Temporary column) | PR review state (`COMMENTED`, `APPROVED`, `CHANGES_REQUESTED`, or `—`) |

---

## Configuration & Synchronization Modes

At the top of `sync_reviews.py`, the `SCAN_ALL_PRS` flag controls scan behavior:

```python
# ==============================================================================
# CONFIGURATION
# ==============================================================================
# Set to True to scan ALL historical Pull Requests in each repository.
# Set to False to use incremental syncing based on latest Excel timestamps.
SCAN_ALL_PRS = False
```

- **`SCAN_ALL_PRS = False` (Default / Production Mode):**  
  Reads `Git_Review_Log.xlsx` from the SVN working directory and identifies the latest review date for each repository. It applies a 10-minute safety buffer and only fetches PRs updated after that timestamp.
- **`SCAN_ALL_PRS = True` (Full Rebuild Mode):**  
  Bypasses Excel timestamp watermarks. Traverses all pages of Pull Requests in repository history. Existing entries in Excel are preserved without duplication thanks to composite key checking.

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

| Parameter Type | Name | Default Value | Description |
|---|---|---|---|
| **String Parameter** | `REPO_LIST` | `https://github.com/org/repo1, https://github.com/org/repo2` | Comma-separated list of GitHub repositories (URLs or slugs). |
| **Password Parameter** | `GITHUB_TOKEN` | `ghp_xxxxxxxxxxxx` | GitHub Personal Access Token (`repo` read access). Concealed by Jenkins. |
| **String Parameter** | `SVN_URL` | `svn://172.16.3.43:3690/company_repo/.../Git_review_log/` | Target SVN directory URL. |
| **String Parameter** | `SVN_USER` | `svnuser` | Username for SVN authentication. |
| **Password Parameter** | `SVN_PASS` | `password123` | Password for SVN authentication. Concealed by Jenkins. |

4. Under **Build Triggers**, select **Build periodically** and configure the schedule:
   ```text
   H 2 * * 1
   ```
   *(Runs every Monday between 02:00 and 03:00 GMT+7)*.
5. Click **Save**.

> **Note on `parameters` in `Jenkinsfile`:**  
> The `parameters { ... }` block is intentionally omitted/commented out inside the `Jenkinsfile`. If declared in code with empty default values, Jenkins would overwrite your UI-saved concealed passwords on every Git push. Defining parameters via the Jenkins UI keeps passwords permanently encrypted on disk.

### 2. Manual vs. Automated Builds

- **Automated Scheduled Runs (Cron):** Jenkins runs in the background and automatically injects the concealed default passwords saved in the job configuration.
- **Manual Runs ("Build with Parameters"):** Jenkins also automatically injects the default inputs and passwords saved in the job configuration.


---

## Project Structure

```text
.
├── Jenkinsfile        # Parameterized CI/CD pipeline definition
├── README.md          # Project documentation
└── sync_reviews.py    # Main synchronization and Excel formatting script
```
