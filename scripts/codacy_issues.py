#!/usr/bin/env python3

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
import re
import subprocess  # nosec B404 - used for fixed git commands only
import sys
from typing import Any, cast
import urllib.parse
import urllib.request

# ================================
# Config
# ================================
BASE = "https://api.codacy.com/api/v3"
BASE_URL = urllib.parse.urlparse(BASE)
BASE_PATH = BASE_URL.path.rstrip("/")  # "/api/v3"


def get_token() -> str:
    """Return the Codacy API token or raise if missing.

    Returns:
        Codacy API token string from the environment.

    Raises:
        ValueError: If CODACY_API_TOKEN is not set.
    """
    token = os.environ.get("CODACY_API_TOKEN")
    if token is None:
        raise ValueError("CODACY_API_TOKEN is not set")
    return token


# ================================
# Utilities
# ================================
LEVELS = ["Error", "High", "Warning", "Info"]


def get_level_priority(level: str | None) -> int | None:
    """Convert a severity level name to a priority number.

    Args:
        level: Severity level string.

    Returns:
        Priority number or None if unknown.
    """
    if level == "Error":
        return 4
    if level == "High":
        return 3
    if level == "Warning":
        return 2
    if level == "Info":
        return 1
    return None


def normalize_provider(value: str) -> str | None:
    """Normalize provider short code.

    Args:
        value: Provider identifier.

    Returns:
        Provider code if valid, otherwise None.
    """
    return value if value in ("gh", "gl", "bb") else None


def assert_valid_segment(name: str, value: str, pattern: re.Pattern[str]) -> str:
    """Validate an identifier segment against a regex.

    Args:
        name: Segment name for error reporting.
        value: Segment value.
        pattern: Compiled regex pattern for allowed values.

    Returns:
        The validated value.

    Raises:
        ValueError: If the value is empty or invalid.
    """
    if (not value) or (pattern.match(value) is None):
        raise ValueError(f"Invalid {name}: {value}")
    return value


def assert_valid_choice(name: str, value: str, choices: list[str]) -> str:
    """Validate that a value is in a list of choices.

    Args:
        name: Parameter name for error reporting.
        value: Input value.
        choices: Allowed values.

    Returns:
        The validated value.

    Raises:
        ValueError: If the value is not allowed.
    """
    if value not in choices:
        raise ValueError(f"Invalid {name}: {value}. Valid values: {', '.join(choices)}")
    return value


def encode_segment(value: str) -> str:
    """URL-encode a path segment."""
    return urllib.parse.quote(value, safe="")


def build_codacy_url(pathname: str, query: dict[str, str] | None = None) -> str:
    """Build a Codacy API URL from a path and query parameters."""
    # Ensure we keep origin and base path
    url = f"{BASE_URL.scheme}://{BASE_URL.netloc}{BASE_PATH}{pathname}"
    if query:
        url = f"{url}?{urllib.parse.urlencode(query)}"
    return url


def assert_codacy_url(url: str) -> str:
    """Ensure the URL targets the Codacy API origin and analysis path.

    Args:
        url: URL to validate.

    Returns:
        The original URL when valid.

    Raises:
        ValueError: If the URL is not within the expected origin/path.
    """
    # Basic safety: must be same origin and start with /api/v3/analysis/
    parsed = urllib.parse.urlparse(url)
    expected_origin = f"{BASE_URL.scheme}://{BASE_URL.netloc}"
    origin = f"{parsed.scheme}://{parsed.netloc}"
    expected_prefix = f"{BASE_PATH}/analysis/"
    if origin != expected_origin or not parsed.path.startswith(expected_prefix):
        raise ValueError(f"Invalid URL: {url}")
    return url


def build_repo_issues_url(provider: str, org: str, repo: str, limit: int) -> str:
    """Build a repository issues API URL."""
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/issues/search",
        query={"limit": str(limit)},
    )


def build_pr_issues_url(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str
) -> str:
    """Build a pull request issues API URL."""
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/pull-requests/{encode_segment(pr)}/issues",
        query={"status": status, "limit": str(limit)},
    )


def get_git_origin_url() -> str | None:
    """Return the git origin URL if available."""
    # git repo check
    try:
        result = subprocess.run(
            ["git", "rev-parse", "--is-inside-work-tree"],
            capture_output=True,
            text=True,
            check=False,
        )  # nosec B603 - fixed git command without user input
        if result.returncode != 0 or not result.stdout.strip():
            return None
        result = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            capture_output=True,
            text=True,
            check=False,
        )  # nosec B603 - fixed git command without user input
        if result.returncode != 0:
            return None
        return result.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return None


@dataclass
class GitRemoteInfo:
    """Parsed git remote information."""

    provider: str
    org: str
    repo: str


def parse_git_remote(url: str) -> GitRemoteInfo | None:
    """Parse a git remote URL into provider/org/repo info."""
    # HTTPS
    m = re.match(r"^https?://([^/]+)/([^/]+)/([^/]+?)(?:\.git)?$", url)
    # SSH
    if not m:
        m = re.match(r"^git@([^:]+):([^/]+)/([^/]+?)(?:\.git)?$", url)

    if not m:
        return None

    host, org, repo = m.group(1), m.group(2), m.group(3)

    def is_same_or_subdomain(hostname: str, base_domain: str) -> bool:
        return hostname == base_domain or hostname.endswith("." + base_domain)

    if is_same_or_subdomain(host, "github.com"):
        provider = "gh"
    elif is_same_or_subdomain(host, "gitlab.com"):
        provider = "gl"
    elif is_same_or_subdomain(host, "bitbucket.org"):
        provider = "bb"
    else:
        provider = "unknown"

    return GitRemoteInfo(provider=provider, org=org, repo=repo)


def fetch_json(
    url: str, method: str = "GET", body: dict[str, Any] | None = None
) -> dict[str, Any]:
    """Fetch JSON from the Codacy API.

    Args:
        url: Codacy API URL.
        method: HTTP method.
        body: Optional JSON body for non-GET requests.

    Returns:
        Parsed JSON dictionary.
    """
    safe_url = assert_codacy_url(url)

    headers = {
        "Accept": "application/json",
        "api-token": get_token(),
    }

    data: bytes | None = None
    if body is not None and method.upper() != "GET":
        payload = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
        headers["Content-Length"] = str(len(payload))
        data = payload

    req = urllib.request.Request(
        safe_url, method=method.upper(), headers=headers, data=data
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as res:  # nosec B310 - validated https origin
            raw = res.read().decode("utf-8", errors="replace")
            try:
                parsed = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise RuntimeError("Invalid JSON response") from exc
            if not isinstance(parsed, dict):
                raise RuntimeError("Invalid JSON response")
            return cast(dict[str, Any], parsed)
    except urllib.error.HTTPError as e:
        # include response body if possible
        try:
            body_text = e.read().decode("utf-8", errors="replace")
        except Exception:
            body_text = ""
        raise RuntimeError(f"HTTP {e.code}: {body_text or str(e)}") from None
    except urllib.error.URLError as e:
        raise RuntimeError(str(e)) from None


# ================================
# API
# ================================
def fetch_repo_issues(provider: str, org: str, repo: str, limit: int) -> dict[str, Any]:
    """Fetch issues for a repository."""
    url = build_repo_issues_url(provider, org, repo, limit)
    return fetch_json(url, method="POST", body={})


def fetch_pr_issues(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str = "all"
) -> dict[str, Any]:
    """Fetch issues for a pull request."""
    url = build_pr_issues_url(provider, org, repo, pr, limit, status)
    return fetch_json(url, method="GET")


# ================================
# AI Output Formatter
# ================================
def format_for_ai(raw_issues: list[dict[str, Any]], min_level: str) -> list[str]:
    """Format raw Codacy issues for AI output.

    Args:
        raw_issues: Issue dictionaries from Codacy API.
        min_level: Minimum severity level to include.

    Returns:
        Formatted issue strings.

    Raises:
        ValueError: If min_level is invalid.
    """
    min_priority = get_level_priority(min_level)
    if min_priority is None:
        raise ValueError(
            f"Invalid min_level: {min_level}. Valid values: {', '.join(LEVELS)}"
        )

    out: list[str] = []

    for item in raw_issues:
        issue = item.get("commitIssue") or item

        pattern_info = issue.get("patternInfo") or {}
        level = pattern_info.get("level")
        prio = get_level_priority(level)
        if prio is None or prio < min_priority:
            continue

        file_path = issue.get("filePath")
        line_no = issue.get("lineNumber")
        rule = pattern_info.get("id")
        category = pattern_info.get("category")
        message = issue.get("message")

        out.append(f"{level} | {file_path}:{line_no} | {rule} | {category} | {message}")

    return out


# ================================
# CLI
# ================================
def parse_args(argv: list[str]) -> argparse.Namespace:
    """Parse command-line arguments."""
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument("org", nargs="?", default=None)
    p.add_argument("repo", nargs="?", default=None)
    p.add_argument("--pr", dest="pr", default=None)
    p.add_argument("--min-level", dest="min_level", default="Info", choices=LEVELS)
    p.add_argument("--provider", dest="provider", default=None)
    p.add_argument("--help", action="help", help="Show this help message and exit")
    return p.parse_args(argv)


def apply_git_defaults(args: argparse.Namespace) -> None:
    """Populate missing org/repo/provider from git origin when possible."""
    if args.org and args.repo:
        return
    origin_url = get_git_origin_url()
    if not origin_url:
        return
    parsed = parse_git_remote(origin_url)
    if not parsed:
        return
    if args.provider is None:
        args.provider = parsed.provider
    if args.org is None:
        args.org = parsed.org
    if args.repo is None:
        args.repo = parsed.repo


def resolve_segments(args: argparse.Namespace) -> tuple[str, str, str | None]:
    """Validate and return org/repo/pr segments.

    Args:
        args: Parsed CLI arguments.

    Returns:
        Tuple of (org, repo, pr).
    """
    segment_pattern = re.compile(r"^[A-Za-z0-9_.-]+$")
    org = assert_valid_segment("org", args.org, segment_pattern)
    repo = assert_valid_segment("repo", args.repo, segment_pattern)
    pr = args.pr
    if pr is not None:
        pr = assert_valid_segment("pr", pr, re.compile(r"^[0-9]+$"))
    return org, repo, pr


def build_payload(
    *,
    pr: str | None,
    org: str,
    repo: str,
    min_level: str,
    issues: list[str],
) -> dict[str, object]:
    """Build the output payload for JSON serialization."""
    return {
        "scope": "pull_request" if pr else "repository",
        "organization": org,
        "repository": repo,
        "pullRequest": pr if pr else None,
        "minLevel": min_level,
        "total": len(issues),
        "issues": issues,
    }


def main() -> int:
    """Run the Codacy issues fetcher."""
    args = parse_args(sys.argv[1:])

    # --- Git auto-detect ---
    apply_git_defaults(args)

    if args.provider is None:
        args.provider = "gh"

    provider = normalize_provider(args.provider)
    if not provider:
        print("Invalid --provider: use gh, gl, or bb", file=sys.stderr)
        return 1

    if not args.org or not args.repo:
        print(
            "Usage:\n"
            "  python codacy_issues.py ORG REPO [--pr NUMBER] [--min-level Error|High|Warning|Info] [--provider gh|gl|bb]",
            file=sys.stderr,
        )
        return 1

    try:
        org, repo, pr = resolve_segments(args)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    status = "all"
    limit = 100

    result = (
        fetch_pr_issues(
            provider=provider, org=org, repo=repo, pr=pr, limit=limit, status=status
        )
        if pr
        else fetch_repo_issues(provider=provider, org=org, repo=repo, limit=limit)
    )

    issues = result.get("data") or []
    try:
        formatted = format_for_ai(issues, args.min_level)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    payload = build_payload(
        pr=pr, org=org, repo=repo, min_level=args.min_level, issues=formatted
    )

    sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(1) from e
