#!/usr/bin/env python3
"""Read-only check: exact live remote branch SHA and clean local checkout."""
import argparse
import json
import subprocess
import sys


def git(repo, *args):
    return subprocess.run(["git", "-C", repo, *args], capture_output=True,
                          text=True, timeout=30, check=True).stdout.strip()


def verify(repo, remote, branch):
    result = {"status": "BLOCKED", "branch": branch, "remote": remote}
    try:
        # Never accept options or arbitrary URLs as the remote argument.
        if not remote or remote.startswith("-") or remote not in git(repo, "remote").splitlines():
            raise ValueError("Remote must be a configured remote name")
        if not branch or branch.startswith("-"):
            raise ValueError("Invalid branch")
        git(repo, "check-ref-format", "refs/heads/" + branch)
        result["local_branch"] = git(repo, "branch", "--show-current")
        result["local_head"] = git(repo, "rev-parse", "HEAD")
        result["dirty"] = bool(git(repo, "status", "--porcelain=v1", "--untracked-files=normal"))
        rows = git(repo, "ls-remote", "--exit-code", "--heads", remote,
                   "refs/heads/" + branch).splitlines()
        matches = [r.split("\t", 1)[0] for r in rows if r.endswith("\trefs/heads/" + branch)]
        if len(matches) != 1:
            raise ValueError("Exact remote ref unavailable")
        result["remote_head"] = matches[0]
        failures = []
        if result["local_branch"] != branch:
            failures.append("WRONG_BRANCH_OR_DETACHED")
        if result["local_head"] != result["remote_head"]:
            failures.append("SHA_MISMATCH")
        if result["dirty"]:
            failures.append("UNCOMMITTED_WORK")
        result["findings"] = failures
        result["status"] = "BLOCKED" if failures else "SYNC_VERIFIED"
    except (subprocess.SubprocessError, OSError, ValueError) as exc:
        # Do not echo potentially sensitive URLs or paths from git stderr.
        result["findings"] = [f"REMOTE_OR_LOCAL_CHECK_FAILED: {type(exc).__name__}"]
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--repo", default=".")
    p.add_argument("--remote", default="origin")
    p.add_argument("--branch", required=True)
    a = p.parse_args()
    result = verify(a.repo, a.remote, a.branch)
    print(json.dumps(result, indent=2))
    return 0 if result["status"] == "SYNC_VERIFIED" else 1


if __name__ == "__main__":
    sys.exit(main())
