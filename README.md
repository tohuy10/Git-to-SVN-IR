# Git-to-SVN-IR (GitHub Review Log Collector & SVN Sync)

An automated tool and Jenkins CI/CD pipeline to aggregate code review activities across multiple GitHub repositories, standardize them into a structured Excel report, and sync the output directly to LARION / Bestarion Subversion (SVN) repositories managed via [LARION SVN](https://svn.larion.com).

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
- **Composite Key Deduplication:** Avoids duplicate rows while preserving multiple identical comments across distinct timestamps and commits.
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
| **I** | File Path / Line (Temporary column) | Target file and line location (e.g., `src/main.py:L12`) |
| **J** | Review State (Temporary column) | PR review state (`COMMENTED`, `APPROVED`, `CHANGES_REQUESTED`) |

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

## Environment Variables

The script expects the following environment variables:

| Variable | Description | Example |
|---|---|---|
| `GITHUB_TOKEN` | GitHub Personal Access Token (`repo` read access) | `ghp_xxxxxxxxxxxx` |
| `REPO_LIST` | Comma-separated list of GitHub repositories | `https://github.com/owner/repo1,https://github.com/owner/repo2` |
| `SVN_URL` | Target SVN directory path URL | `svn://svn.larion.com/company_repo/Project_Engineering/6_Review/Git_review_log/` |
| `SVN_USER` | SVN authentication username | `svnuser` |
| `SVN_PASS` | SVN authentication password | `password123` |

---

## Jenkins Pipeline Setup

### 1. Credentials Configuration
Navigate to **Manage Jenkins** → **Credentials** → **System** → **Global credentials**:
- **`github-token`**: Type **Secret text** (enter your GitHub personal access token).
- **`svn-credentials`**: Type **Username with password** (enter your SVN credentials).

### 2. Job Configuration
1. Create a **New Item** → Select **Pipeline**.
2. Point the pipeline script to this Git repository or paste the provided `Jenkinsfile`.
3. Run the build once via **Build Now** to register parameters.
4. Subsequent runs will use **Build with Parameters**, allowing you to supply `REPO_LIST` and `SVN_URL` directly from the web UI.

---

## Project Structure

```text
.
├── Jenkinsfile        # Parameterized CI/CD pipeline definition
├── README.md          # Project documentation
└── sync_reviews.py    # Main synchronization and Excel formatting script
```
