from __future__ import annotations

import base64
import json
import re
from dataclasses import dataclass
from datetime import date, datetime, timezone
from typing import Any, Dict, Optional
from urllib.parse import quote

import requests
import streamlit as st


class GitHubDraftSaveError(RuntimeError):
    """Raised when a draft cannot be saved to GitHub."""


@dataclass
class GitHubSaveResult:
    path: str
    html_url: str
    commit_sha: str


def slugify(value: str) -> str:
    """Make a filename-safe slug from a presenter name or session title."""
    value = str(value or "").strip().lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "untitled"


def build_draft_filename(presenter_name: str, session_title: str) -> str:
    """Use today's date + presenter + session title for easy retrieval."""
    today = date.today().isoformat()
    presenter_slug = slugify(presenter_name) or "unknown-presenter"
    session_slug = slugify(session_title) or "untitled-session"
    return f"{today}_{presenter_slug}_{session_slug}.json"


def _read_github_config() -> Dict[str, str]:
    """Read GitHub backup configuration from Streamlit secrets."""
    try:
        raw = st.secrets.get("github", {})
    except Exception:
        raw = {}

    return {
        "token": str(raw.get("token", "")).strip(),
        "repo": str(raw.get("repo", "")).strip(),
        "branch": str(raw.get("branch", "main")).strip() or "main",
        "base_path": str(raw.get("base_path", "drafts")).strip().strip("/") or "drafts",
    }


def github_backup_is_configured() -> bool:
    cfg = _read_github_config()
    return bool(cfg["token"] and cfg["repo"] and "/" in cfg["repo"])


def github_config_status_message() -> str:
    cfg = _read_github_config()
    missing = []
    if not cfg["token"]:
        missing.append("github.token")
    if not cfg["repo"] or "/" not in cfg["repo"]:
        missing.append("github.repo")
    if missing:
        return "Missing Streamlit secrets: " + ", ".join(missing)
    return f"Configured to save drafts to {cfg['repo']} on branch {cfg['branch']}."


def make_draft_payload(
    deck: Dict[str, Any],
    presenter_name: str,
    session_title: str,
    app_version: str,
) -> Dict[str, Any]:
    return {
        "presenter_name": presenter_name.strip(),
        "session_title": session_title.strip(),
        "saved_date": date.today().isoformat(),
        "saved_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "app_version": app_version,
        "deck": deck,
    }


def save_draft_to_github(
    deck: Dict[str, Any],
    presenter_name: str,
    session_title: str,
    app_version: str,
) -> GitHubSaveResult:
    """Create or update a JSON draft in the configured GitHub repo."""
    cfg = _read_github_config()

    if not github_backup_is_configured():
        raise GitHubDraftSaveError(github_config_status_message())

    filename = build_draft_filename(presenter_name, session_title)
    path = f"{cfg['base_path']}/{filename}"
    api_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"

    payload = make_draft_payload(
        deck=deck,
        presenter_name=presenter_name,
        session_title=session_title,
        app_version=app_version,
    )
    json_text = json.dumps(payload, indent=2, ensure_ascii=False)
    encoded_content = base64.b64encode(json_text.encode("utf-8")).decode("utf-8")

    headers = {
        "Authorization": f"Bearer {cfg['token']}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    existing_sha: Optional[str] = None
    existing = requests.get(
        api_url,
        headers=headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )
    if existing.status_code == 200:
        existing_sha = existing.json().get("sha")
    elif existing.status_code != 404:
        raise GitHubDraftSaveError(
            f"Could not check existing GitHub file ({existing.status_code}): {existing.text}"
        )

    data: Dict[str, Any] = {
        "message": f"Save journal club draft: {filename}",
        "content": encoded_content,
        "branch": cfg["branch"],
    }
    if existing_sha:
        data["sha"] = existing_sha

    response = requests.put(api_url, headers=headers, json=data, timeout=30)
    if response.status_code not in (200, 201):
        raise GitHubDraftSaveError(
            f"GitHub save failed ({response.status_code}): {response.text}"
        )

    result = response.json()
    content = result.get("content", {}) or {}
    commit = result.get("commit", {}) or {}
    return GitHubSaveResult(
        path=content.get("path", path),
        html_url=content.get("html_url", ""),
        commit_sha=commit.get("sha", ""),
    )
