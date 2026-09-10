#!/usr/bin/env python3
import os
import re
import sys
import subprocess
from datetime import datetime, timedelta, timezone
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

EXCEL_FILE_NAME = "Git_Review_Log.xlsx"
SHEET_NAME = "Review Logs"
commit_cache = {}


def normalize_repo_slug(raw: str) -> str:
    cleaned = raw.strip()
    match = re.search(r"github\.com[/:]([\w-]+/[\w.-]+?)(?:\.git)?$", cleaned)
    return match.group(1) if match else cleaned


def run_cmd(cmd, cwd=None):
    result = subprocess.run(cmd, cwd=cwd, text=True, capture_output=True)
    if result.returncode != 0:
        print(f"[CMD ERROR] {' '.join(cmd)}\nSTDERR: {result.stderr.strip()}", file=sys.stderr)
        raise subprocess.CalledProcessError(result.returncode, cmd, result.stdout, result.stderr)
    return result.stdout.strip()


def sync_svn(work_dir: str, svn_url: str, user: str, password: str):
    auth_args = ["--non-interactive", "--no-auth-cache"]
    if user:
        auth_args.extend(["--username", user, "--password", password])

    if not os.path.exists(os.path.join(work_dir, ".svn")):
        if os.path.exists(work_dir):
            import shutil
            shutil.rmtree(work_dir)
        cmd = ["svn", "checkout", svn_url, work_dir] + auth_args
        print(f"Executing: svn checkout {svn_url} {work_dir}")
        run_cmd(cmd)
    else:
        cmd = ["svn", "update", work_dir] + auth_args
        print("Executing: svn update")
        run_cmd(cmd)


def commit_to_svn(work_dir: str, user: str, password: str):
    auth_args = ["--non-interactive", "--no-auth-cache"]
    if user:
        auth_args.extend(["--username", user, "--password", password])

    status = run_cmd(["svn", "status", EXCEL_FILE_NAME], cwd=work_dir)
    if not status:
        print("No changes detected in Excel log. Skipping SVN commit.")
        return

    if status.startswith("?"):
        run_cmd(["svn", "add", "--force", EXCEL_FILE_NAME], cwd=work_dir)

    commit_cmd = ["svn", "commit", "-m", "[Auto-Sync] Update Git review logs via Jenkins", EXCEL_FILE_NAME] + auth_args
    output = run_cmd(commit_cmd, cwd=work_dir)
    print(output)


def get_commit_details(repo: str, sha: str, headers: dict):
    if not sha:
        return {}
    if sha in commit_cache:
        return commit_cache[sha]

    url = f"https://api.github.com/repos/{repo}/commits/{sha}"
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code == 200:
        data = resp.json()
        commit_cache[sha] = data
        return data
    return {}


def process_repository(repo: str, token: str, since_dt: datetime):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    all_repo_records = []

    # PRs sorted by creation date ascending
    prs_url = f"https://api.github.com/repos/{repo}/pulls?state=all&sort=created&direction=asc&per_page=50"
    resp = requests.get(prs_url, headers=headers, timeout=15)
    if resp.status_code != 200:
        print(f"Failed to fetch PRs for {repo}: {resp.status_code} - {resp.text}")
        return all_repo_records

    prs = resp.json()

    for pr in prs:
        pr_number = pr["number"]
        default_dev = pr.get("user", {}).get("login", "")
        pr_messages = []

        # 1. Main PR Reviews (/reviews)
        reviews_url = f"https://api.github.com/repos/{repo}/pulls/{pr_number}/reviews?per_page=50"
        r_resp = requests.get(reviews_url, headers=headers, timeout=15)
        if r_resp.status_code == 200:
            for rev in r_resp.json():
                body = (rev.get("body") or "").strip()
                if not body:
                    continue

                submitted_at_str = rev.get("submitted_at")
                if submitted_at_str:
                    sub_dt = datetime.fromisoformat(submitted_at_str.replace("Z", "+00:00"))
                    if sub_dt < since_dt:
                        continue
                else:
                    sub_dt = datetime.min.replace(tzinfo=timezone.utc)
                    submitted_at_str = ""

                commit_sha = rev.get("commit_id") or pr.get("head", {}).get("sha", "")
                commit_info = get_commit_details(repo, commit_sha, headers)

                dev_name = (
                    commit_info.get("commit", {}).get("author", {}).get("name")
                    or default_dev
                )
                commit_msg = commit_info.get("commit", {}).get("message", "")

                pr_messages.append({
                    "raw_dt": sub_dt,
                    "repo_name": repo,
                    "commit_id": commit_sha,
                    "dev_name": dev_name,
                    "created_at": submitted_at_str.replace("T", " ").replace("Z", ""),
                    "commit_content": commit_msg,
                    "reviewer_name": rev.get("user", {}).get("login", ""),
                    "reviewer_comment": body,
                    "review_type": "PR Review",
                    "file_path_line": "—",
                    "review_state": rev.get("state", ""),
                })

        # 2. Line-Specific Comments (/comments)
        comments_url = f"https://api.github.com/repos/{repo}/pulls/{pr_number}/comments?per_page=100"
        c_resp = requests.get(comments_url, headers=headers, timeout=15)
        if c_resp.status_code == 200:
            for c in c_resp.json():
                created_at_str = c.get("created_at")
                if created_at_str:
                    c_dt = datetime.fromisoformat(created_at_str.replace("Z", "+00:00"))
                    if c_dt < since_dt:
                        continue
                else:
                    c_dt = datetime.min.replace(tzinfo=timezone.utc)
                    created_at_str = ""

                commit_sha = c.get("commit_id") or pr.get("head", {}).get("sha", "")
                commit_info = get_commit_details(repo, commit_sha, headers)

                dev_name = (
                    commit_info.get("commit", {}).get("author", {}).get("name")
                    or default_dev
                )
                commit_msg = commit_info.get("commit", {}).get("message", "")

                path = c.get("path", "")
                line = c.get("line")
                start_line = c.get("start_line")

                if start_line and line and start_line != line:
                    loc = f"{path}:L{start_line}-L{line}"
                elif line:
                    loc = f"{path}:L{line}"
                else:
                    loc = path

                pr_messages.append({
                    "raw_dt": c_dt,
                    "repo_name": repo,
                    "commit_id": commit_sha,
                    "dev_name": dev_name,
                    "created_at": created_at_str.replace("T", " ").replace("Z", ""),
                    "commit_content": commit_msg,
                    "reviewer_name": c.get("user", {}).get("login", ""),
                    "reviewer_comment": c.get("body", ""),
                    "review_type": "Line Specific",
                    "file_path_line": loc,
                    "review_state": "—",
                })

        # Intertwine reviews and comments by time
        pr_messages.sort(key=lambda item: item["raw_dt"])
        all_repo_records.extend(pr_messages)

    return all_repo_records


def update_excel(file_path: str, records: list):
    headers = [
        "Repository", "Commit ID", "Dev name", "Created at", "Commit content",
        "Reviewer name", "Reviewer comment", "Review Type", "File Path / Line", "Review State"
    ]
    existing_keys = set()

    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 7:
                # Key: Repo (0) | Commit ID (1) | Reviewer (5) | Created At (3) | Comment (6)
                comment_text = str(row[6] or "").strip()
                key = f"{row[0]}|{row[1]}|{row[5]}|{row[3]}|{comment_text}"
                existing_keys.add(key)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(headers)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_align = Alignment(horizontal="center", vertical="center")

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
        ws.row_dimensions[1].height = 25

    data_align = Alignment(vertical="top", wrap_text=True)
    added_count = 0

    for rec in records:
        clean_comment = rec["reviewer_comment"].strip()
        key = f"{rec['repo_name']}|{rec['commit_id']}|{rec['reviewer_name']}|{rec['created_at']}|{clean_comment}"
        if key in existing_keys:
            continue
        existing_keys.add(key)

        row_data = [
            rec["repo_name"],
            rec["commit_id"],
            rec["dev_name"],
            rec["created_at"],
            rec["commit_content"],
            rec["reviewer_name"],
            rec["reviewer_comment"].replace("\r\n", "\n"),
            rec["review_type"],
            rec["file_path_line"],
            rec["review_state"],
        ]

        ws.append(row_data)
        current_row = ws.max_row
        for col_idx in range(1, len(row_data) + 1):
            ws.cell(row=current_row, column=col_idx).alignment = data_align
        added_count += 1

    col_widths = {
        "A": 28, "B": 22, "C": 16, "D": 20, "E": 30,
        "F": 16, "G": 45, "H": 15, "I": 22, "J": 14
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(file_path)
    print(f"Excel sync complete: {added_count} records added.")


def main():
    github_token = os.getenv("GITHUB_TOKEN")
    repo_list_env = os.getenv("REPO_LIST")
    svn_url = os.getenv("SVN_URL")
    svn_user = os.getenv("SVN_USER", "")
    svn_pass = os.getenv("SVN_PASS", "")



    if not github_token:
        print("[ERROR] GITHUB_TOKEN environment variable not set.", file=sys.stderr)
        sys.exit(1)
    if not repo_list_env:
        print("[ERROR] REPO_LIST environment variable not set.", file=sys.stderr)
        sys.exit(1)
    if not svn_url:
        print("[ERROR] SVN_URL environment variable not set.", file=sys.stderr)
        sys.exit(1)

    work_dir = "svn_workdir"

    print("==> Step 1: Syncing SVN Working Directory...")
    sync_svn(work_dir, svn_url, svn_user, svn_pass)

    print("==> Step 2: Fetching GitHub Review Logs...")
    since_dt = datetime.now(timezone.utc) - timedelta(days=7)
    all_records = []

    for repo_raw in repo_list_env.split(","):
        repo = normalize_repo_slug(repo_raw)
        if not repo:
            continue
        print(f"Processing repository: {repo}")
        records = process_repository(repo, github_token, since_dt)
        all_records.extend(records)

    print(f"==> Total review records collected: {len(all_records)}")

    excel_path = os.path.join(work_dir, EXCEL_FILE_NAME)
    print(f"==> Step 3: Updating Excel file at {excel_path}...")
    update_excel(excel_path, all_records)

    print("==> Step 4: Committing updated Excel to SVN...")
    commit_to_svn(work_dir, svn_user, svn_pass)

    print("[SUCCESS] Process completed successfully.")


if __name__ == "__main__":
    main()