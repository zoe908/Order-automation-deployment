"""
업체 등록 시 vendors/configs/*.json 파일을 GitHub 레포에 자동 커밋합니다.
앱이 재시작돼도 레포에서 복원되어 데이터가 유지됩니다.
"""

import base64
import requests


def push_file(token: str, repo: str, file_path: str, content_bytes: bytes, commit_message: str):
    """
    GitHub API로 파일을 생성 또는 업데이트합니다.

    token: GitHub Personal Access Token
    repo:  "owner/repo-name" 형식 (예: "zoe908/Order-automation-deployment")
    file_path: 레포 내 경로 (예: "vendors/configs/hanssem.json")
    content_bytes: 파일 내용 (bytes)
    commit_message: 커밋 메시지
    """
    url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }

    # 기존 파일이면 SHA 필요 (없으면 새 파일 생성)
    existing = requests.get(url, headers=headers)
    sha = existing.json().get("sha") if existing.status_code == 200 else None

    payload = {
        "message": commit_message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(url, headers=headers, json=payload)
    resp.raise_for_status()
