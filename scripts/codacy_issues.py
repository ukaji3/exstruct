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
TOKEN = os.environ.get("CODACY_API_TOKEN")

if TOKEN is None:
    print("CODACY_API_TOKEN is not set", file=sys.stderr)
    sys.exit(1)
TOKEN_STR = TOKEN


# ================================
# Utilities
# ================================
LEVELS = ["Error", "High", "Warning", "Info"]


def get_level_priority(level: str | None) -> int | None:
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
    return value if value in ("gh", "gl", "bb") else None


def assert_valid_segment(name: str, value: str, pattern: re.Pattern[str]) -> str:
    if (not value) or (pattern.match(value) is None):
        print(f"Invalid {name}: {value}", file=sys.stderr)
        sys.exit(1)
    return value


def assert_valid_choice(name: str, value: str, choices: list[str]) -> str:
    if value not in choices:
        print(
            f"Invalid {name}: {value}. Valid values: {', '.join(choices)}",
            file=sys.stderr,
        )
        sys.exit(1)
    return value


def encode_segment(value: str) -> str:
    return urllib.parse.quote(value, safe="")


def build_codacy_url(pathname: str, query: dict[str, str] | None = None) -> str:
    # Ensure we keep origin and base path
    url = f"{BASE_URL.scheme}://{BASE_URL.netloc}{BASE_PATH}{pathname}"
    if query:
        url = f"{url}?{urllib.parse.urlencode(query)}"
    return url


def assert_codacy_url(url: str) -> str:
    # Basic safety: must be same origin and start with /api/v3/analysis/
    parsed = urllib.parse.urlparse(url)
    expected_origin = f"{BASE_URL.scheme}://{BASE_URL.netloc}"
    origin = f"{parsed.scheme}://{parsed.netloc}"
    expected_prefix = f"{BASE_PATH}/analysis/"
    if origin != expected_origin or not parsed.path.startswith(expected_prefix):
        print(f"Invalid URL: {url}", file=sys.stderr)
        sys.exit(1)
    return url


def build_repo_issues_url(provider: str, org: str, repo: str, limit: int) -> str:
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/issues/search",
        query={"limit": str(limit)},
    )


def build_pr_issues_url(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str
) -> str:
    return build_codacy_url(
        f"/analysis/organizations/{encode_segment(provider)}/{encode_segment(org)}"
        f"/repositories/{encode_segment(repo)}/pull-requests/{encode_segment(pr)}/issues",
        query={"status": status, "limit": str(limit)},
    )


def get_git_origin_url() -> str | None:
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
    except Exception:
        return None


@dataclass
class GitRemoteInfo:
    provider: str
    org: str
    repo: str


def parse_git_remote(url: str) -> GitRemoteInfo | None:
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
    safe_url = assert_codacy_url(url)

    headers = {
        "Accept": "application/json",
        "api-token": TOKEN_STR,
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
            status = getattr(res, "status", 0) or 0
            if status < 200 or status >= 300:
                raise RuntimeError(f"HTTP {status}: {raw}")
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
    url = build_repo_issues_url(provider, org, repo, limit)
    return fetch_json(url, method="POST", body={})


def fetch_pr_issues(
    provider: str, org: str, repo: str, pr: str, limit: int, status: str = "all"
) -> dict[str, Any]:
    url = build_pr_issues_url(provider, org, repo, pr, limit, status)
    return fetch_json(url, method="GET")


# ================================
# AI Output Formatter
# ================================
def format_for_ai(raw_issues: list[dict[str, Any]], min_level: str) -> list[str]:
    min_priority = get_level_priority(min_level)
    if min_priority is None:
        print(
            f"Invalid --min-level: {min_level}. Valid values: {', '.join(LEVELS)}",
            file=sys.stderr,
        )
        sys.exit(1)

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
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument("org", nargs="?", default=None)
    p.add_argument("repo", nargs="?", default=None)
    p.add_argument("--pr", dest="pr", default=None)
    p.add_argument("--min-level", dest="min_level", default="Info", choices=LEVELS)
    p.add_argument("--provider", dest="provider", default=None)
    p.add_argument("--help", action="help", help="Show this help message and exit")
    return p.parse_args(argv)


def main() -> int:
    args = parse_args(sys.argv[1:])

    # --- Git auto-detect ---
    if not args.org or not args.repo:
        origin_url = get_git_origin_url()
        if origin_url:
            parsed = parse_git_remote(origin_url)
            if parsed:
                if args.provider is None:
                    args.provider = parsed.provider
                if args.org is None:
                    args.org = parsed.org
                if args.repo is None:
                    args.repo = parsed.repo

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

    segment_pattern = re.compile(r"^[A-Za-z0-9_.-]+$")
    org = assert_valid_segment("org", args.org, segment_pattern)
    repo = assert_valid_segment("repo", args.repo, segment_pattern)
    pr = args.pr
    if pr is not None:
        pr = assert_valid_segment("pr", pr, re.compile(r"^[0-9]+$"))

    status = assert_valid_choice("status", "all", ["all", "open", "closed"])
    limit = 100

    result = (
        fetch_pr_issues(
            provider=provider, org=org, repo=repo, pr=pr, limit=limit, status=status
        )
        if pr
        else fetch_repo_issues(provider=provider, org=org, repo=repo, limit=limit)
    )

    issues = result.get("data") or []
    formatted = format_for_ai(issues, args.min_level)

    payload = {
        "scope": "pull_request" if pr else "repository",
        "organization": org,
        "repository": repo,
        "pullRequest": pr if pr else None,
        "minLevel": args.min_level,
        "total": len(formatted),
        "issues": formatted,
    }

    sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(str(e), file=sys.stderr)
        raise SystemExit(1) from e
