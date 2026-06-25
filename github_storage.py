from __future__ import annotations

import base64
import hmac
import json
import re
import uuid
from dataclasses import dataclass
from datetime import date, datetime, timezone
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import streamlit as st


class GitHubDraftSaveError(RuntimeError):
    """Raised when a draft cannot be saved to GitHub."""


class GitHubDraftLoadError(RuntimeError):
    """Raised when a draft cannot be listed or loaded from GitHub."""


class GitHubDraftDeleteError(RuntimeError):
    """Raised when a draft or article cannot be deleted from GitHub."""


@dataclass
class GitHubResult:
    path: str
    html_url: str
    commit_sha: str


def slugify(value: str) -> str:
    """Make a filename-safe slug from a presenter name, session title, or filename."""
    value = str(value or "").strip().lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "untitled"


def normalize_archive_id(archive_id: str | None) -> str:
    """Keep archive IDs short and filename-safe."""
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "", str(archive_id or "")).lower()
    return cleaned[:24]


def generate_archive_id() -> str:
    """Create a short unique ID used to tie one draft JSON to its article PDF."""
    return uuid.uuid4().hex[:12]


def build_draft_filename(presenter_name: str, session_title: str, archive_id: str | None = None) -> str:
    """Use date + presenter + session title + optional unique archive ID."""
    today = date.today().isoformat()
    presenter_slug = slugify(presenter_name) or "unknown-presenter"
    session_slug = slugify(session_title) or "untitled-session"
    archive_suffix = normalize_archive_id(archive_id)
    if archive_suffix:
        return f"{today}_{presenter_slug}_{session_slug}_{archive_suffix}.json"
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
        "delete_passcode": str(raw.get("delete_passcode", "")).strip(),
    }


def _github_headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
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
    return "Storage Archive Activated and Ready for Use"


def github_delete_is_configured() -> bool:
    """Return True only when normal GitHub storage and a deletion passcode are configured."""
    cfg = _read_github_config()
    return github_backup_is_configured() and bool(cfg.get("delete_passcode"))


def github_delete_status_message() -> str:
    """Human-readable status for delete controls."""
    if not github_backup_is_configured():
        return github_config_status_message()
    cfg = _read_github_config()
    if not cfg.get("delete_passcode"):
        return "Deletion is disabled. Add github.delete_passcode to Streamlit secrets to enable it."
    return "Deletion controls are enabled."


def github_delete_passcode_is_valid(passcode: str) -> bool:
    """Check the entered delete passcode without exposing the saved secret."""
    cfg = _read_github_config()
    expected = str(cfg.get("delete_passcode", ""))
    provided = str(passcode or "")
    if not expected or not provided:
        return False
    return hmac.compare_digest(provided, expected)


def make_draft_payload(
    deck: Dict[str, Any],
    presenter_name: str,
    session_title: str,
    app_version: str,
    article: Optional[Dict[str, Any]] = None,
    archive_id: str | None = None,
) -> Dict[str, Any]:
    archive_id = normalize_archive_id(archive_id) or generate_archive_id()
    article_payload = dict(article or {})
    if article_payload:
        article_payload.setdefault("archive_id", archive_id)

    return {
        "archive_id": archive_id,
        "presenter_name": presenter_name.strip(),
        "session_title": session_title.strip(),
        "saved_date": date.today().isoformat(),
        "saved_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "app_version": app_version,
        "article": article_payload,
        "deck": deck,
    }


def save_draft_to_github(
    deck: Dict[str, Any],
    presenter_name: str,
    session_title: str,
    app_version: str,
    article: Optional[Dict[str, Any]] = None,
    archive_id: str | None = None,
) -> GitHubResult:
    """Create or update a JSON draft in the configured GitHub repo."""
    cfg = _read_github_config()

    if not github_backup_is_configured():
        raise GitHubDraftSaveError(github_config_status_message())

    archive_id = normalize_archive_id(archive_id) or generate_archive_id()
    filename = build_draft_filename(presenter_name, session_title, archive_id=archive_id)
    path = f"{cfg['base_path']}/{filename}"
    api_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"

    payload = make_draft_payload(
        deck=deck,
        presenter_name=presenter_name,
        session_title=session_title,
        app_version=app_version,
        article=article,
        archive_id=archive_id,
    )
    json_text = json.dumps(payload, indent=2, ensure_ascii=False)
    encoded_content = base64.b64encode(json_text.encode("utf-8")).decode("utf-8")

    headers = _github_headers(cfg["token"])

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
    return GitHubResult(
        path=content.get("path", path),
        html_url=content.get("html_url", ""),
        commit_sha=commit.get("sha", ""),
    )


def save_article_to_github(
    article_bytes: bytes,
    original_filename: str,
    presenter_name: str,
    session_title: str,
    archive_id: str | None = None,
) -> GitHubResult:
    """Create or update the uploaded article PDF in the configured GitHub repo.

    The archive ID makes the article path unique for each saved journal club.
    The original uploaded filename is preserved in JSON metadata, but the stored
    path uses a stable `..._<archive_id>_article.pdf` name so replacing the PDF
    updates the same archive record instead of creating orphaned files.
    """
    cfg = _read_github_config()

    if not github_backup_is_configured():
        raise GitHubDraftSaveError(github_config_status_message())

    archive_id = normalize_archive_id(archive_id) or generate_archive_id()
    base_filename = build_draft_filename(presenter_name, session_title, archive_id=archive_id).replace(".json", "")
    extension = original_filename.rsplit(".", 1)[-1].lower() if "." in original_filename else "pdf"
    extension = slugify(extension).replace("-", "") or "pdf"

    filename = f"{base_filename}_article.{extension}"
    path = f"{cfg['base_path']}/articles/{filename}"

    api_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"

    encoded_content = base64.b64encode(article_bytes).decode("utf-8")
    headers = _github_headers(cfg["token"])

    existing_sha = None
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
            f"Could not check existing GitHub article ({existing.status_code}): {existing.text}"
        )

    data: Dict[str, Any] = {
        "message": f"Save journal club article: {filename}",
        "content": encoded_content,
        "branch": cfg["branch"],
    }

    if existing_sha:
        data["sha"] = existing_sha

    response = requests.put(api_url, headers=headers, json=data, timeout=30)

    if response.status_code not in (200, 201):
        raise GitHubDraftSaveError(
            f"GitHub article save failed ({response.status_code}): {response.text}"
        )

    result = response.json()
    content = result.get("content", {}) or {}
    commit = result.get("commit", {}) or {}

    return GitHubResult(
        path=content.get("path", path),
        html_url=content.get("html_url", ""),
        commit_sha=commit.get("sha", ""),
    )


def list_drafts_from_github(presenter_name: str = "") -> List[Dict[str, Any]]:
    """List saved JSON drafts from the configured GitHub drafts folder."""
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubDraftLoadError(github_config_status_message())

    base_path = cfg["base_path"]
    api_path = quote(base_path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"
    headers = _github_headers(cfg["token"])

    response = requests.get(
        api_url,
        headers=headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if response.status_code == 404:
        return []

    if response.status_code != 200:
        raise GitHubDraftLoadError(
            f"Could not list GitHub drafts ({response.status_code}): {response.text}"
        )

    items = response.json()
    if not isinstance(items, list):
        raise GitHubDraftLoadError("GitHub drafts path is not a folder. Check github.base_path in secrets.")

    presenter_slug = slugify(presenter_name) if str(presenter_name or "").strip() else ""
    drafts: List[Dict[str, Any]] = []
    for item in items:
        if item.get("type") != "file":
            continue
        name = str(item.get("name", ""))
        if not name.endswith(".json"):
            continue
        if presenter_slug and presenter_slug not in name:
            continue
        drafts.append(
            {
                "name": name,
                "path": item.get("path", ""),
                "sha": item.get("sha", ""),
                "html_url": item.get("html_url", ""),
                "size": item.get("size", 0),
            }
        )

    drafts.sort(key=lambda item: str(item.get("name", "")), reverse=True)
    return drafts


def load_file_bytes_from_github(path: str) -> bytes:
    """Load any saved file from the configured GitHub repo as bytes."""
    cfg = _read_github_config()

    if not github_backup_is_configured():
        raise GitHubDraftLoadError(github_config_status_message())

    clean_path = str(path or "").strip().lstrip("/")
    if not clean_path:
        raise GitHubDraftLoadError("No GitHub file path was provided.")

    api_path = quote(clean_path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"

    headers = _github_headers(cfg["token"])

    response = requests.get(
        api_url,
        headers=headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if response.status_code != 200:
        raise GitHubDraftLoadError(
            f"Could not load GitHub file ({response.status_code}): {response.text}"
        )

    try:
        payload = response.json()
    except Exception:
        payload = {}

    encoded_content = str(payload.get("content", "") or "").replace("\n", "")
    encoding = str(payload.get("encoding", "") or "").lower()

    if encoded_content and encoding in {"base64", ""}:
        try:
            return base64.b64decode(encoded_content)
        except Exception as exc:
            raise GitHubDraftLoadError(f"GitHub file content could not be decoded: {exc}") from exc

    raw_headers = _github_headers(cfg["token"])
    raw_headers["Accept"] = "application/vnd.github.raw"
    raw_response = requests.get(
        api_url,
        headers=raw_headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if raw_response.status_code == 200 and raw_response.content:
        return raw_response.content

    raise GitHubDraftLoadError(
        f"GitHub file did not contain downloadable content. Raw download also failed "
        f"({raw_response.status_code}): {raw_response.text}"
    )


def load_draft_from_github(path: str) -> Dict[str, Any]:
    """Load and decode one JSON draft from GitHub."""
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubDraftLoadError(github_config_status_message())

    clean_path = str(path or "").strip().lstrip("/")
    if not clean_path:
        raise GitHubDraftLoadError("No GitHub draft path was provided.")

    api_path = quote(clean_path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"
    headers = _github_headers(cfg["token"])

    response = requests.get(
        api_url,
        headers=headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )
    if response.status_code != 200:
        raise GitHubDraftLoadError(
            f"Could not load GitHub draft ({response.status_code}): {response.text}"
        )

    payload = response.json()
    encoded = str(payload.get("content", "")).replace("\n", "")
    if not encoded:
        raise GitHubDraftLoadError("GitHub draft did not contain file content.")

    try:
        decoded = base64.b64decode(encoded).decode("utf-8")
        data = json.loads(decoded)
    except Exception as exc:
        raise GitHubDraftLoadError(f"GitHub draft is not valid JSON: {exc}") from exc

    return data


def delete_file_from_github(path: str, commit_message: Optional[str] = None, missing_ok: bool = False) -> GitHubResult:
    """Delete one file from the configured GitHub repo."""
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubDraftDeleteError(github_config_status_message())

    clean_path = str(path or "").strip().lstrip("/")
    if not clean_path:
        raise GitHubDraftDeleteError("No GitHub file path was provided for deletion.")

    api_path = quote(clean_path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"
    headers = _github_headers(cfg["token"])

    existing = requests.get(
        api_url,
        headers=headers,
        params={"ref": cfg["branch"]},
        timeout=30,
    )

    if existing.status_code == 404 and missing_ok:
        return GitHubResult(path=clean_path, html_url="", commit_sha="")

    if existing.status_code != 200:
        raise GitHubDraftDeleteError(
            f"Could not find GitHub file to delete ({existing.status_code}): {existing.text}"
        )

    payload = existing.json()
    sha = payload.get("sha")
    html_url = payload.get("html_url", "")
    if not sha:
        raise GitHubDraftDeleteError("GitHub file did not include a SHA, so it cannot be deleted.")

    data: Dict[str, Any] = {
        "message": commit_message or f"Delete journal club file: {clean_path}",
        "sha": sha,
        "branch": cfg["branch"],
    }

    response = requests.delete(api_url, headers=headers, json=data, timeout=30)
    if response.status_code not in (200, 204):
        raise GitHubDraftDeleteError(
            f"GitHub delete failed ({response.status_code}): {response.text}"
        )

    try:
        result = response.json() if response.text else {}
    except Exception:
        result = {}
    commit = result.get("commit", {}) or {}

    return GitHubResult(path=clean_path, html_url=html_url, commit_sha=commit.get("sha", ""))


def find_drafts_referencing_article(article_path: str, exclude_draft_path: str = "") -> List[str]:
    """Find other JSON drafts that point to the same saved article path."""
    target = str(article_path or "").strip().lstrip("/")
    exclude = str(exclude_draft_path or "").strip().lstrip("/")
    if not target:
        return []

    references: List[str] = []
    for draft in list_drafts_from_github(""):
        draft_path = str(draft.get("path", "") or "").strip().lstrip("/")
        if not draft_path or draft_path == exclude:
            continue
        try:
            loaded = load_draft_from_github(draft_path)
        except Exception:
            continue
        other_article_path = ""
        if isinstance(loaded, dict) and isinstance(loaded.get("article"), dict):
            other_article_path = str(loaded.get("article", {}).get("path", "") or "").strip().lstrip("/")
        if other_article_path == target:
            references.append(draft_path)
    return references


def delete_draft_and_article_from_github(draft_path: str, delete_article: bool = True) -> Dict[str, List[str]]:
    """Delete a saved JSON draft and optionally its associated article PDF.

    Safety check: if another saved draft points to the same article path, the
    article is not deleted. This protects older archives created before unique
    archive IDs were added.
    """
    clean_draft_path = str(draft_path or "").strip().lstrip("/")
    if not clean_draft_path:
        raise GitHubDraftDeleteError("No GitHub draft path was provided for deletion.")

    deleted: List[str] = []
    warnings: List[str] = []

    loaded = load_draft_from_github(clean_draft_path)
    article_path = ""
    if isinstance(loaded, dict) and isinstance(loaded.get("article"), dict):
        article_path = str(loaded.get("article", {}).get("path", "") or "").strip().lstrip("/")

    if delete_article:
        if article_path and article_path != clean_draft_path:
            other_refs = find_drafts_referencing_article(article_path, exclude_draft_path=clean_draft_path)
            if other_refs:
                warnings.append(
                    "Article PDF was not deleted because another saved draft still references it: "
                    + ", ".join(other_refs[:5])
                    + (" ..." if len(other_refs) > 5 else "")
                )
            else:
                result = delete_file_from_github(
                    article_path,
                    commit_message=f"Delete journal club article for {clean_draft_path}",
                    missing_ok=True,
                )
                if result.commit_sha:
                    deleted.append(result.path)
                else:
                    warnings.append("Associated article PDF was already missing from GitHub.")
        else:
            warnings.append("No associated article PDF path was recorded in this draft.")

    result = delete_file_from_github(
        clean_draft_path,
        commit_message=f"Delete journal club draft: {clean_draft_path}",
        missing_ok=False,
    )
    deleted.append(result.path)

    return {"deleted": deleted, "warnings": warnings}
