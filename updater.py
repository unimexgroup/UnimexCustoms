"""
GitHub-Releases self-updater for the frozen (PyInstaller) exes.
=================================================================
Standard-library only -- no new runtime dependency. This module is only
meaningful when running as a packaged .exe; the caller guards on ``sys.frozen``
and imports it lazily so a dev run never touches it.

Design principles:
  * Never raises to the caller. Offline, DNS failure, GitHub 5xx, rate-limit
    (403), a changed JSON shape -- everything falls through to a one-line note
    and ``return False`` so the user's real work is never blocked.
  * Only called after a run that had PROBLEMS (skipped / failed / unrecognized
    files). A clean run never checks -- no wasted time when the tool works.
  * Releases are tagged ``customs-vX.Y.Z`` so the tag prefix stays explicit and
    a future second tool in this repo can version independently. The updater
    filters to its own prefix and picks the highest semver among those.
  * The running .exe is locked on Windows and cannot overwrite itself, so we
    download a sidecar and hand the swap to a tiny detached .cmd that waits for
    this process to exit, moves the new file into place, and relaunches.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

OWNER = "unimexgroup"
REPO = "UnimexCustoms"
RELEASES_URL = f"https://api.github.com/repos/{OWNER}/{REPO}/releases"

HTTP_TIMEOUT = 8  # seconds; keeps offline users from hanging
UA = f"{REPO}-updater"  # GitHub requires a User-Agent header on API requests


# ---------------------------------------------------------------------------
# Semantic version parsing / comparison
# ---------------------------------------------------------------------------
def _parse(ver: str) -> tuple[int, int, int]:
    """Turn '1.2.3' (or 'v1.2', 'customs-v1.2.3') into a comparable (1, 2, 3)."""
    ver = ver.strip()
    if "-v" in ver:  # strip a 'customs-v' prefix if present
        ver = ver.split("-v", 1)[1]
    ver = ver.lstrip("vV")
    nums: list[int] = []
    for part in ver.split(".")[:3]:
        digits = "".join(ch for ch in part if ch.isdigit())
        nums.append(int(digits) if digits else 0)
    while len(nums) < 3:
        nums.append(0)
    return nums[0], nums[1], nums[2]


# ---------------------------------------------------------------------------
# GitHub query
# ---------------------------------------------------------------------------
def _fetch_json(url: str):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": UA, "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _latest_release_for(prefix: str):
    """Return (version_str, release_dict) for the highest published release whose
    tag starts with '<prefix>-v', or (None, None). Drafts and pre-releases are
    skipped (drafts aren't visible anonymously anyway)."""
    releases = _fetch_json(RELEASES_URL)
    tag_pre = f"{prefix}-v"
    best_ver: tuple[int, int, int] | None = None
    best_rel = None
    for rel in releases:
        if rel.get("draft") or rel.get("prerelease"):
            continue
        tag = rel.get("tag_name", "")
        if not tag.startswith(tag_pre):
            continue
        pv = _parse(tag[len(tag_pre):])
        if best_ver is None or pv > best_ver:
            best_ver, best_rel = pv, rel
    if best_rel is None:
        return None, None
    return best_rel["tag_name"][len(tag_pre):], best_rel


def _find_asset(rel, asset_name: str):
    for a in rel.get("assets", []):
        if a.get("name") == asset_name:
            return a
    return None


# ---------------------------------------------------------------------------
# Download + integrity check
# ---------------------------------------------------------------------------
def _download(url: str, dest: Path, expected_size: int) -> bool:
    """Stream the asset to `dest`. Verify the byte count against the size the
    API reported (GitHub exposes no per-asset checksum) to catch truncation."""
    req = urllib.request.Request(
        url, headers={"User-Agent": UA, "Accept": "application/octet-stream"}
    )
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp, open(dest, "wb") as f:
        while True:
            chunk = resp.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
    if expected_size and dest.stat().st_size != expected_size:
        try:
            dest.unlink()
        except Exception:
            pass
        return False
    return True


# ---------------------------------------------------------------------------
# Windows locked-exe self-replace
# ---------------------------------------------------------------------------
def _spawn_swap_helper(exe: Path, new_file: Path) -> None:
    """Write a throwaway .cmd in %TEMP% that swaps the downloaded file over the
    running .exe once we release our lock on it, relaunches it, and deletes
    itself. Spawned fully detached so it outlives us.

    Robustness notes (learned the hard way):
      * We do NOT wait on our PID with ``tasklist``/``find`` -- a Windows exe is
        locked against replacement while it runs, so we simply retry ``move``:
        it fails while we're still alive and succeeds the instant we exit. This
        also avoids ``find``/``tasklist`` being shadowed by unrelated tools on
        PATH.
      * We sleep with a full-path ``ping`` rather than ``timeout``, because
        ``timeout`` aborts immediately in a detached process with no console.
      * ``move``, ``start`` and ``del`` are cmd built-ins, so nothing here
        depends on PATH resolution. All paths are baked in and quoted so spaces
        in the install path can't break the script.
    """
    helper = Path(tempfile.gettempdir()) / f"_update_{exe.stem}_{os.getpid()}.cmd"
    ping = r"%SystemRoot%\System32\ping.exe"
    script = f"""@echo off
setlocal
set "OLD={exe}"
set "NEW={new_file}"

set /a tries=0
:swap
move /Y "%NEW%" "%OLD%" >NUL 2>&1
if not errorlevel 1 goto relaunch
set /a tries+=1
if %tries% geq 60 goto done
{ping} -n 2 127.0.0.1 >NUL
goto swap

:relaunch
start "" "%OLD%"

:done
(goto) 2>nul & del "%~f0"
"""
    helper.write_text(script, encoding="ascii")

    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    CREATE_NO_WINDOW = 0x08000000
    subprocess.Popen(
        ["cmd", "/c", str(helper)],
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW,
        close_fds=True,
    )


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def check_and_update(tool_prefix: str, asset_name: str, current_version: str) -> bool:
    """Check GitHub Releases for a newer build of this tool and, if found,
    download it and schedule an in-place swap + relaunch.

    Returns True only when an update was installed and a relaunch is scheduled
    -- in that case the caller must exit WITHOUT running its work or pausing,
    because a fresh instance is taking over. Returns False for 'already current',
    'offline', 'not frozen', or any error (caller keeps running as normal).
    """
    if not getattr(sys, "frozen", False):
        return False  # dev run -- nothing to update

    exe = Path(sys.executable)
    try:
        latest, rel = _latest_release_for(tool_prefix)
        if not latest:
            return False
        if _parse(latest) <= _parse(current_version):
            print(f"[update] You have the latest version (v{current_version}).")
            return False

        asset = _find_asset(rel, asset_name)
        if not asset:
            print(f"[update] v{latest} is available but no {asset_name} was "
                  f"attached to the release; keeping current version.")
            return False

        print(f"[update] New version v{latest} available "
              f"(you have v{current_version}). Downloading...")
        new_file = exe.with_suffix(exe.suffix + ".new")  # same folder = same volume
        if not _download(asset["browser_download_url"], new_file, asset.get("size", 0)):
            print("[update] Download was incomplete; keeping current version.")
            return False

        print(f"[update] Installing v{latest} and restarting...")
        _spawn_swap_helper(exe, new_file)
        return True
    except Exception as e:
        print(f"[update] Update check skipped ({type(e).__name__}). "
              f"Continuing on the current version.")
        return False
