#!/usr/bin/env python3
"""SVN Excel Diff — Web UI

A local web server that provides:
  - Directory browser for selecting SVN working copies
  - Changed file list (svn status)
  - Cell-level Excel diff with IDE-style side-by-side view

Usage:
    python server.py [--port 5000]

Then open http://localhost:5000 in your browser.
"""

import os
import platform
import subprocess
import sys
import tempfile
import webbrowser

from flask import Flask, jsonify, request

# Import diff engine from the same directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from svn_excel_diff import read_excel_to_rows, build_unified_diff, get_svn_base, get_svn_command

app = Flask(__name__)


# ---------------------------------------------------------------------------
# API: Directory browsing
# ---------------------------------------------------------------------------

@app.route("/api/browse")
def api_browse():
    """List subdirectories and files under a given path.
    Query params: path (default: user home)
    """
    req_path = request.args.get("path", "").strip()
    if not req_path:
        req_path = os.path.expanduser("~")

    # Normalize
    req_path = os.path.normpath(os.path.expanduser(req_path))

    if not os.path.isdir(req_path):
        return jsonify({"error": f"Not a directory: {req_path}"}), 400

    entries = []
    try:
        for name in sorted(os.listdir(req_path), key=str.lower):
            full = os.path.join(req_path, name)
            if name.startswith("."):
                continue
            is_dir = os.path.isdir(full)
            entries.append({
                "name": name,
                "path": full,
                "is_dir": is_dir,
            })
    except PermissionError:
        return jsonify({"error": f"Permission denied: {req_path}"}), 403

    # Detect if this directory is under SVN control
    # (either has .svn itself, or is inside an SVN working copy)
    is_svn_root = os.path.isdir(os.path.join(req_path, ".svn"))
    is_svn = is_svn_root
    if not is_svn:
        try:
            r = subprocess.run(
                [get_svn_command(), "info", req_path],
                capture_output=True, text=True, timeout=5,
            )
            is_svn = r.returncode == 0
        except Exception:
            pass

    return jsonify({
        "path": req_path,
        "parent": os.path.dirname(req_path),
        "is_svn": is_svn,
        "entries": entries,
        "sep": os.sep,
    })


@app.route("/api/drives")
def api_drives():
    """List available drives (Windows) or root + home (Unix)."""
    if platform.system() == "Windows":
        import ctypes
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = []
        for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            if bitmask & 1:
                drives.append({"name": f"{letter}:\\", "path": f"{letter}:\\"})
            bitmask >>= 1
        return jsonify({"drives": drives})
    else:
        home = os.path.expanduser("~")
        return jsonify({"drives": [
            {"name": "/", "path": "/"},
            {"name": "Home", "path": home},
        ]})


# ---------------------------------------------------------------------------
# API: SVN operations
# ---------------------------------------------------------------------------

@app.route("/api/svn/status")
def api_svn_status():
    """Get svn status for a directory. Returns list of changed files."""
    dir_path = request.args.get("path", "").strip()
    if not dir_path or not os.path.isdir(dir_path):
        return jsonify({"error": "Invalid path"}), 400

    try:
        result = subprocess.run(
            [get_svn_command(), "status", dir_path],
            capture_output=True, text=True, timeout=30,
        )
    except FileNotFoundError:
        return jsonify({"error": "svn command not found. Please install SVN or add it to PATH."}), 500
    except subprocess.TimeoutExpired:
        return jsonify({"error": "svn status timed out"}), 500

    files = []
    for line in result.stdout.strip().split("\n"):
        if not line.strip():
            continue
        status = line[0]
        filepath = line[8:].strip() if len(line) > 8 else line[1:].strip()
        if status == "?":
            continue
        ext = os.path.splitext(filepath)[1].lower()
        files.append({
            "status": status,
            "path": filepath,
            "name": os.path.basename(filepath),
            "rel_path": os.path.relpath(filepath, dir_path),
            "type": "excel" if ext in (".xls", ".xlsx") else "text",
        })

    return jsonify({
        "path": dir_path,
        "files": files,
        "count": len(files),
    })


@app.route("/api/svn/diff")
def api_svn_diff():
    """Get cell-level diff for a single file.
    Query params: file (absolute path)
    """
    filepath = request.args.get("file", "").strip()
    if not filepath or not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 400

    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xls", ".xlsx"):
        base_tmp = get_svn_base(filepath)
        if base_tmp is None:
            return jsonify({"error": "Could not get SVN base version"}), 500

        try:
            base_data = read_excel_to_rows(base_tmp)
            work_data = read_excel_to_rows(filepath)
            udiff = build_unified_diff(base_data, work_data)

            # Convert to JSON-serializable format
            result = {}
            for sheet_name, sdata in udiff.items():
                rows = []
                for r in sdata["rows"]:
                    rows.append({
                        "type": r["type"],
                        "key": r["key"],
                        "base": r["base"],
                        "work": r["work"],
                        "changed_cols": list(r["changed_cols"]),
                    })
                result[sheet_name] = {
                    "headers": sdata["headers"],
                    "base_headers": sdata["base_headers"],
                    "rows": rows,
                }
            return jsonify({
                "file": filepath,
                "name": os.path.basename(filepath),
                "type": "excel",
                "sheets": result,
            })
        finally:
            os.unlink(base_tmp)
    else:
        # Text file — use svn diff
        result = subprocess.run(
            [get_svn_command(), "diff", filepath],
            capture_output=True, text=True,
        )
        return jsonify({
            "file": filepath,
            "name": os.path.basename(filepath),
            "type": "text",
            "diff": result.stdout,
        })


@app.route("/api/svn/log")
def api_svn_log():
    """Get SVN commit history.
    Query params: path, limit (default 50)
    """
    dir_path = request.args.get("path", "").strip()
    limit = request.args.get("limit", "50")
    if not dir_path:
        return jsonify({"error": "Invalid path"}), 400

    username = request.args.get("username", "").strip()
    password = request.args.get("password", "").strip()

    cmd = [get_svn_command(), "log", "-l", str(limit), "--xml", "--non-interactive"]
    if username and password:
        cmd += ["--username", username, "--password", password]
    cmd.append(dir_path)

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except FileNotFoundError:
        return jsonify({"error": "svn command not found. Please install SVN or add it to PATH."}), 500
    except subprocess.TimeoutExpired:
        return jsonify({"error": "svn log timed out"}), 500

    if result.returncode != 0:
        return jsonify({"error": result.stderr.strip() or "svn log failed"}), 500

    # Parse XML output
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(result.stdout)
    except ET.ParseError:
        return jsonify({"error": "Failed to parse svn log"}), 500

    entries = []
    for entry in root.findall("logentry"):
        rev = entry.get("revision")
        author = entry.findtext("author", "")
        date = entry.findtext("date", "")
        msg = entry.findtext("msg", "")
        entries.append({
            "revision": rev,
            "author": author,
            "date": date,
            "message": msg,
        })

    return jsonify({"path": dir_path, "entries": entries})


@app.route("/api/svn/log-detail")
def api_svn_log_detail():
    """Get changed files for a specific revision.
    Query params: path, revision
    """
    dir_path = request.args.get("path", "").strip()
    rev = request.args.get("revision", "").strip()
    if not dir_path or not rev:
        return jsonify({"error": "path and revision required"}), 400

    username = request.args.get("username", "").strip()
    password = request.args.get("password", "").strip()

    cmd = [get_svn_command(), "log", "-r", rev, "-v", "--xml", "--non-interactive"]
    if username and password:
        cmd += ["--username", username, "--password", password]
    cmd.append(dir_path)

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except Exception:
        return jsonify({"error": "svn log failed"}), 500

    if result.returncode != 0:
        return jsonify({"error": result.stderr.strip() or "svn log failed"}), 500

    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(result.stdout)
    except ET.ParseError:
        return jsonify({"error": "Failed to parse svn log"}), 500

    entry = root.find("logentry")
    if entry is None:
        return jsonify({"error": "Revision not found"}), 404

    files = []
    paths_el = entry.find("paths")
    if paths_el is not None:
        for p in paths_el.findall("path"):
            action = p.get("action", "?")
            filepath = p.text or ""
            files.append({
                "action": action,
                "path": filepath,
                "name": os.path.basename(filepath),
            })

    return jsonify({
        "revision": rev,
        "author": entry.findtext("author", ""),
        "date": entry.findtext("date", ""),
        "message": entry.findtext("msg", ""),
        "files": files,
    })


def _get_repo_root(wc_path):
    """Get the SVN repository root URL from a working copy path."""
    try:
        result = subprocess.run(
            [get_svn_command(), "info", "--xml", wc_path],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode == 0:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(result.stdout)
            repo_root = root.findtext(".//repository/root")
            return repo_root
    except Exception:
        pass
    return None


def _svn_cat_rev(full_url, rev, username=None, password=None):
    """Get file content at a specific revision via full SVN URL. Returns temp file path or None."""
    try:
        cmd = [get_svn_command(), "cat", "-r", str(rev), "--non-interactive"]
        if username and password:
            cmd += ["--username", username, "--password", password]
        cmd.append(full_url)
        print(f"[DEBUG] svn cat cmd: svn cat -r {rev} {full_url}", flush=True)
        result = subprocess.run(cmd, capture_output=True, check=True, timeout=30)
        ext = os.path.splitext(full_url)[1]
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as f:
            f.write(result.stdout)
            return f.name
    except subprocess.CalledProcessError as e:
        print(f"[DEBUG] svn cat failed: {e.stderr.decode() if e.stderr else e}", flush=True)
        return None
    except Exception as e:
        print(f"[DEBUG] svn cat exception: {e}", flush=True)
        return None


@app.route("/api/svn/history-diff")
def api_svn_history_diff():
    """Get diff for a file at a specific revision (rev vs rev-1).
    Query params: path (working copy path), file (repo path from log), revision
    """
    wc_path = request.args.get("path", "").strip()
    file_repo_path = request.args.get("file", "").strip()
    rev = request.args.get("revision", "").strip()
    username = request.args.get("username", "").strip()
    password = request.args.get("password", "").strip()
    if not wc_path or not file_repo_path or not rev:
        return jsonify({"error": "path, file, and revision required"}), 400

    rev_int = int(rev)
    prev_rev = rev_int - 1
    ext = os.path.splitext(file_repo_path)[1].lower()
    fname = os.path.basename(file_repo_path)

    # Construct full SVN URL: repo_root + repo_path
    repo_root = _get_repo_root(wc_path)
    if not repo_root:
        return jsonify({"error": "Could not determine SVN repository root"}), 500
    full_url = repo_root + file_repo_path
    print(f"[DEBUG] history-diff: repo_root={repo_root}, file_repo_path={file_repo_path}, full_url={full_url}, rev={rev_int}, prev={prev_rev}", flush=True)

    if ext in (".xls", ".xlsx"):
        old_tmp = _svn_cat_rev(full_url, prev_rev, username, password)
        new_tmp = _svn_cat_rev(full_url, rev_int, username, password)

        if old_tmp is None and new_tmp is None:
            return jsonify({"error": f"Could not get file at either revision. URL: {full_url}"}), 500

        try:
            old_data = read_excel_to_rows(old_tmp) if old_tmp else {}
            new_data = read_excel_to_rows(new_tmp) if new_tmp else {}
            udiff = build_unified_diff(old_data, new_data)

            result = {}
            for sheet_name, sdata in udiff.items():
                rows = []
                for r in sdata["rows"]:
                    rows.append({
                        "type": r["type"],
                        "key": r["key"],
                        "base": r["base"],
                        "work": r["work"],
                        "changed_cols": list(r["changed_cols"]),
                    })
                result[sheet_name] = {
                    "headers": sdata["headers"],
                    "base_headers": sdata["base_headers"],
                    "rows": rows,
                }
            return jsonify({
                "file": file_repo_path,
                "name": fname,
                "type": "excel",
                "revision": rev,
                "sheets": result,
            })
        finally:
            if old_tmp:
                os.unlink(old_tmp)
            if new_tmp:
                os.unlink(new_tmp)
    else:
        # Text file
        try:
            cmd = [get_svn_command(), "diff", "-c", rev, "--non-interactive"]
            if username and password:
                cmd += ["--username", username, "--password", password]
            cmd.append(full_url)
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=30,
            )
            return jsonify({
                "file": file_repo_path,
                "name": fname,
                "type": "text",
                "revision": rev,
                "diff": result.stdout,
            })
        except Exception as e:
            return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Frontend — served as a single page
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return FRONTEND_HTML


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(description="SVN Excel Diff Web UI")
    parser.add_argument("--port", type=int, default=9527)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    if not args.no_browser:
        import threading
        threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{args.port}")).start()

    print(f"Starting SVN Diff Web UI at http://localhost:{args.port}")
    app.run(host="0.0.0.0", port=args.port, debug=False)


# ---------------------------------------------------------------------------
# Frontend HTML (embedded for single-file distribution)
# ---------------------------------------------------------------------------

FRONTEND_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SVN Diff Viewer</title>
<style>
:root {
  --bg: #1e1e1e;
  --bg-surface: #252526;
  --bg-header: #2d2d2d;
  --bg-sidebar: #252526;
  --text: #d4d4d4;
  --text-dim: #808080;
  --text-bright: #ffffff;
  --border: #404040;
  --accent: #569cd6;
  --accent-hover: #6cb0f0;
  --added-bg: #1e3a1e;
  --added-text: #3fb950;
  --deleted-bg: #3a1e1e;
  --deleted-text: #f85149;
  --modified-bg: #3a3a1e;
  --modified-cell: #d29922;
  --modified-cell-bg: rgba(210,153,34,0.15);
  --hover-bg: #2a2d2e;
  --selected-bg: #37373d;
  --scrollbar-thumb: #555;
  --input-bg: #3c3c3c;
  --input-border: #555;
}
* { margin:0; padding:0; box-sizing:border-box; }
body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  font-size: 13px;
  background: var(--bg);
  color: var(--text);
  height: 100vh;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

/* ---- Top bar ---- */
.topbar {
  background: var(--bg-header);
  border-bottom: 1px solid var(--border);
  padding: 8px 16px;
  display: flex;
  align-items: center;
  gap: 12px;
  flex-shrink: 0;
}
.topbar .logo { font-size:14px; font-weight:700; color:var(--accent); white-space:nowrap; }
.path-display {
  flex: 1;
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
}
.path-display .current-path {
  background: var(--input-bg);
  border: 1px solid var(--input-border);
  border-radius: 4px;
  padding: 5px 12px;
  font-family: 'Menlo','Consolas',monospace;
  font-size: 12px;
  color: var(--text);
  cursor: pointer;
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  transition: border-color 0.15s;
  max-width: 700px;
}
.path-display .current-path:hover { border-color: var(--accent); }
.path-display .current-path.empty { color: var(--text-dim); font-style: italic; }
.btn {
  background: var(--accent);
  color: #fff;
  border: none;
  padding: 5px 14px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
  transition: background 0.15s;
}
.btn:hover { background: var(--accent-hover); }
.btn-outline {
  background: transparent;
  border: 1px solid var(--border);
  color: var(--text);
}
.btn-outline:hover { border-color: var(--accent); color: var(--accent); background: rgba(86,156,214,0.08);}
.btn-sm { padding: 3px 10px; font-size: 11px; }

/* ---- Main layout ---- */
.main { display:flex; flex:1; overflow:hidden; }

/* ---- Sidebar ---- */
.sidebar {
  width: 300px; min-width:200px; max-width:500px;
  background: var(--bg-sidebar);
  border-right: 1px solid var(--border);
  display: flex; flex-direction: column; flex-shrink: 0;
  resize: horizontal; overflow: auto;
}
.sidebar-header {
  padding:10px 14px; font-size:11px; font-weight:600;
  text-transform:uppercase; letter-spacing:0.5px; color:var(--text-dim);
  border-bottom:1px solid var(--border);
  display:flex; align-items:center; justify-content:space-between;
}
.sidebar-header .count {
  background:var(--accent); color:#fff;
  padding:1px 6px; border-radius:8px; font-size:10px;
}
.file-list { flex:1; overflow-y:auto; padding:4px 0; }
.file-item {
  padding:6px 14px; cursor:pointer;
  display:flex; align-items:center; gap:8px;
  transition:background 0.1s;
  border-left:3px solid transparent;
}
.file-item:hover { background:var(--hover-bg); }
.file-item.active { background:var(--selected-bg); border-left-color:var(--accent); }
.file-item .status-badge {
  font-size:10px; font-weight:700; padding:1px 5px; border-radius:3px; flex-shrink:0;
}
.file-item .status-badge.M { background:var(--modified-bg); color:var(--modified-cell); }
.file-item .status-badge.A { background:var(--added-bg); color:var(--added-text); }
.file-item .status-badge.D { background:var(--deleted-bg); color:var(--deleted-text); }
.file-item .file-name { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:13px; }
.file-item .file-path { font-size:11px; color:var(--text-dim); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.file-item-info { flex:1; overflow:hidden; }

/* ---- Content ---- */
.content { flex:1; display:flex; flex-direction:column; overflow:hidden; }
.content-placeholder {
  flex:1; display:flex; align-items:center; justify-content:center;
  color:var(--text-dim); font-size:14px; flex-direction:column; gap:12px;
}
.content-placeholder .hint { font-size:12px; color:var(--text-dim); opacity:0.6; }

/* ---- Sheet tabs ---- */
.sheet-tabs { display:flex; background:var(--bg-header); border-bottom:1px solid var(--border); overflow-x:auto; flex-shrink:0; }
.sheet-tab {
  padding:7px 16px; cursor:pointer; border-bottom:2px solid transparent;
  color:var(--text-dim); white-space:nowrap; font-size:12px; user-select:none; transition:all 0.15s;
}
.sheet-tab:hover { background:var(--hover-bg); color:var(--text); }
.sheet-tab.active { color:var(--accent); border-bottom-color:var(--accent); }
.sheet-tab .tab-stats { font-size:10px; margin-left:6px; padding:1px 5px; border-radius:3px; background:var(--modified-bg); color:var(--modified-cell); }
.sheet-tab .tab-stats.clean { background:var(--bg-surface); color:var(--text-dim); }

/* ---- Filter bar ---- */
.filter-bar { background:var(--bg-surface); padding:6px 14px; border-bottom:1px solid var(--border); display:flex; gap:8px; align-items:center; flex-shrink:0; }
.filter-label { color:var(--text-dim); font-size:11px; }
.filter-btn { background:var(--bg-header); border:1px solid var(--border); color:var(--text-dim); padding:2px 10px; border-radius:3px; cursor:pointer; font-size:11px; font-family:inherit; transition:all 0.15s; }
.filter-btn:hover { border-color:var(--accent); color:var(--text); }
.filter-btn.active { border-color:var(--accent); color:var(--accent); background:rgba(86,156,214,0.1); }
.stats-summary { margin-left:auto; display:flex; gap:10px; }
.stat-chip { font-size:11px; font-weight:600; padding:2px 8px; border-radius:3px; }
.stat-chip.added { background:var(--added-bg); color:var(--added-text); }
.stat-chip.deleted { background:var(--deleted-bg); color:var(--deleted-text); }
.stat-chip.modified { background:var(--modified-bg); color:var(--modified-cell); }

/* ---- Diff split layout ---- */
.diff-labels { display:flex; flex-shrink:0; border-bottom:1px solid var(--border); }
.diff-label {
  flex:1; text-align:center; padding:6px 8px;
  font-size:12px; font-weight:600; color:var(--accent); letter-spacing:1px;
  background:var(--bg-header);
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
}
.diff-label.left { border-right:3px solid var(--accent); }
.diff-scroll { flex:1; overflow-y:auto; display:flex; }
.diff-half { flex:1; overflow-x:auto; min-width:0; }
.diff-half.left { border-right:3px solid var(--accent); }

/* ---- Diff table ---- */
.diff-table { width:100%; border-collapse:collapse; table-layout:auto; font-family:'Menlo','Consolas','Courier New',monospace; font-size:12px; }
.diff-table th { position:sticky; top:0; background:var(--bg-header); color:var(--text-dim); font-weight:600; font-size:11px; text-align:left; padding:5px 8px; border-bottom:2px solid var(--border); border-right:1px solid var(--border); white-space:nowrap; z-index:10; }
.diff-table th.side-label { text-align:center; font-size:12px; color:var(--accent); letter-spacing:1px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; }
.diff-table tbody tr { height: 24px; }
.diff-table td { padding:3px 8px; border-right:1px solid rgba(255,255,255,0.04); border-bottom:1px solid rgba(255,255,255,0.04); white-space:nowrap; user-select:text; cursor:text; }
.diff-table td.ln { color:var(--text-dim); text-align:right; user-select:none; width:36px; min-width:36px; padding-right:6px; font-size:11px; }
.diff-table td.gutter { width:4px; min-width:4px; padding:0; border-right:none; }
tr.row-equal td { background:transparent; }
tr.row-added td { background:var(--added-bg); }
tr.row-added td.empty-side { background:var(--bg-surface); color:var(--text-dim); }
tr.row-deleted td { background:var(--deleted-bg); }
tr.row-deleted td.empty-side { background:var(--bg-surface); color:var(--text-dim); }
tr.row-modified td { background:transparent; }
tr.row-modified td.cell-changed { background:var(--modified-cell-bg); }
tr.row-added td.gutter { background:var(--added-text); }
tr.row-deleted td.gutter { background:var(--deleted-text); }
tr.row-modified td.gutter { background:var(--modified-cell); }
.old-val { background:rgba(248,81,73,0.25); color:var(--deleted-text); padding:0 2px; border-radius:2px; }
.new-val { background:rgba(46,160,67,0.25); color:var(--added-text); padding:0 2px; border-radius:2px; }

/* Collapsed rows */
tr.row-collapsed td.collapse-cell {
  text-align:center; padding:6px 8px; cursor:pointer;
  background:var(--bg-surface); color:var(--accent);
  font-size:12px; user-select:none; border-bottom:1px solid var(--border);
  transition:background 0.15s;
}
tr.row-collapsed td.collapse-cell:hover { background:var(--hover-bg); }

/* Text diff */
.text-diff { flex:1; overflow:auto; padding:16px; font-family:'Menlo','Consolas',monospace; font-size:12px; white-space:pre-wrap; line-height:1.6; }
.text-diff .line-add { color:var(--added-text); }
.text-diff .line-del { color:var(--deleted-text); }
.text-diff .line-hunk { color:var(--accent); font-weight:600; }

/* ---- Browse modal ---- */
.modal-overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); z-index:1000; align-items:center; justify-content:center; }
.modal-overlay.open { display:flex; }
.modal {
  background:var(--bg-surface); border:1px solid var(--border); border-radius:8px;
  width:620px; max-height:75vh; display:flex; flex-direction:column;
}
.modal-header {
  padding:12px 16px; border-bottom:1px solid var(--border);
  display:flex; align-items:center; justify-content:space-between; font-weight:600;
}
.modal-header .close-btn { background:none; border:none; color:var(--text-dim); font-size:18px; cursor:pointer; padding:0 4px; }
.modal-header .close-btn:hover { color:var(--text); }

/* Breadcrumb */
.breadcrumb-bar {
  padding:8px 16px; border-bottom:1px solid var(--border);
  display:flex; align-items:center; gap:2px; overflow-x:auto; flex-shrink:0;
}
.breadcrumb-item {
  color:var(--accent); cursor:pointer; padding:2px 6px; border-radius:3px;
  white-space:nowrap; font-size:12px; transition:background 0.15s;
}
.breadcrumb-item:hover { background:var(--hover-bg); }
.breadcrumb-item.current { color:var(--text-bright); cursor:default; font-weight:600; }
.breadcrumb-item.current:hover { background:transparent; }
.breadcrumb-sep { color:var(--text-dim); font-size:11px; user-select:none; }

/* Quick access */
.quick-access {
  padding:10px 16px; border-bottom:1px solid var(--border);
  display:flex; gap:6px; flex-wrap:wrap;
}
.quick-btn {
  background:var(--bg-header); border:1px solid var(--border); color:var(--text);
  padding:4px 12px; border-radius:4px; cursor:pointer; font-size:12px;
  display:flex; align-items:center; gap:5px; transition:all 0.15s;
}
.quick-btn:hover { border-color:var(--accent); color:var(--accent); }
.quick-btn .qicon { font-size:14px; }
.quick-section-label { color:var(--text-dim); font-size:11px; width:100%; margin-bottom:2px; }

/* Directory list */
.modal-body { flex:1; overflow-y:auto; }
.dir-item {
  padding:7px 16px; cursor:pointer; display:flex; align-items:center; gap:10px;
  font-size:13px; transition:background 0.1s; border-left:3px solid transparent;
}
.dir-item:hover { background:var(--hover-bg); }
.dir-item .dir-icon { font-size:16px; flex-shrink:0; width:20px; text-align:center; }
.dir-item .dir-name { flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.dir-item .svn-tag {
  font-size:9px; font-weight:700; padding:1px 5px; border-radius:3px;
  background:var(--added-bg); color:var(--added-text); flex-shrink:0;
}
.dir-item .changes-tag {
  font-size:9px; font-weight:700; padding:1px 5px; border-radius:3px;
  background:var(--modified-bg); color:var(--modified-cell); flex-shrink:0;
}

.modal-footer {
  padding:10px 16px; border-top:1px solid var(--border);
  display:flex; align-items:center; gap:10px;
}
.modal-footer .selected-label {
  flex:1; font-size:12px; color:var(--text-dim);
  font-family:monospace; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
}

/* ---- History modal ---- */
.history-modal { width:750px; max-height:80vh; }
.log-list { flex:1; overflow-y:auto; }
.log-item {
  padding:10px 16px; cursor:pointer; border-bottom:1px solid var(--border);
  transition:background 0.1s;
}
.log-item:hover { background:var(--hover-bg); }
.log-item.active { background:var(--selected-bg); border-left:3px solid var(--accent); }
.log-rev { color:var(--accent); font-weight:600; font-size:12px; }
.log-meta { color:var(--text-dim); font-size:11px; margin-left:8px; }
.log-msg { margin-top:3px; font-size:12px; color:var(--text); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }

.history-layout { display:flex; flex:1; overflow:hidden; }
.history-sidebar { width:350px; border-right:1px solid var(--border); display:flex; flex-direction:column; overflow:hidden; }
.history-detail { flex:1; display:flex; flex-direction:column; overflow:hidden; }
.history-detail-header {
  padding:8px 14px; background:var(--bg-header); border-bottom:1px solid var(--border);
  font-size:12px; color:var(--text-dim); flex-shrink:0;
}
.history-file-list { flex:1; overflow-y:auto; }
.history-file-item {
  padding:6px 14px; cursor:pointer; display:flex; align-items:center; gap:8px;
  transition:background 0.1s; font-size:13px;
}
.history-file-item:hover { background:var(--hover-bg); }
.history-file-item .action-badge {
  font-size:10px; font-weight:700; padding:1px 5px; border-radius:3px; flex-shrink:0;
}
.history-file-item .action-badge.A { background:var(--added-bg); color:var(--added-text); }
.history-file-item .action-badge.D { background:var(--deleted-bg); color:var(--deleted-text); }
.history-file-item .action-badge.M { background:var(--modified-bg); color:var(--modified-cell); }

/* Loading / spinner */
.loading { display:flex; align-items:center; justify-content:center; padding:40px; color:var(--text-dim); gap:10px; }
.spinner { width:18px; height:18px; border:2px solid var(--border); border-top-color:var(--accent); border-radius:50%; animation:spin 0.6s linear infinite; }
@keyframes spin { to { transform:rotate(360deg); } }

/* Scrollbar */
::-webkit-scrollbar { width:8px; height:8px; }
::-webkit-scrollbar-track { background:var(--bg); }
::-webkit-scrollbar-thumb { background:var(--scrollbar-thumb); border-radius:4px; }
</style>
</head>
<body>

<!-- Top bar -->
<div class="topbar">
  <div class="logo">SVN Diff</div>
  <div class="path-display">
    <div class="current-path empty" id="pathDisplay" onclick="openBrowse()">Click to select SVN directory...</div>
    <button class="btn" id="scanBtn" onclick="scanChanges()">Scan Changes</button>
    <button class="btn-outline btn" id="refreshBtn" onclick="refreshChanges()" style="display:none;" title="Refresh">&#8635; Refresh</button>
    <button class="btn-outline btn" onclick="openHistory()">History</button>
  </div>
</div>

<!-- Main -->
<div class="main">
  <div class="sidebar">
    <div class="sidebar-header">
      <span>Changed Files</span>
      <span class="count" id="fileCount">0</span>
    </div>
    <div class="file-list" id="fileList">
      <div class="content-placeholder" style="padding:30px 14px;font-size:12px;">
        <span>Select a directory to start</span>
      </div>
    </div>
  </div>
  <div class="content" id="contentArea">
    <div class="content-placeholder">
      <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14,2 14,8 20,8"/><line x1="9" y1="15" x2="15" y2="15" style="stroke:var(--deleted-text)"/><line x1="9" y1="15" x2="15" y2="15" transform="translate(0,3)" style="stroke:var(--added-text)"/></svg>
      <span>Select a file to view its diff</span>
      <span class="hint">Click the path bar above to browse and select an SVN directory</span>
    </div>
  </div>
</div>

<!-- Browse modal -->
<div class="modal-overlay" id="browseModal">
  <div class="modal">
    <div class="modal-header">
      <span>Select SVN Directory</span>
      <button class="close-btn" onclick="closeBrowse()">&times;</button>
    </div>
    <div class="breadcrumb-bar" id="breadcrumbBar"></div>
    <div class="quick-access" id="quickAccess"></div>
    <div class="modal-body" id="modalBody"></div>
    <div class="modal-footer">
      <span class="selected-label" id="selectedLabel">No directory selected</span>
      <button class="btn-outline btn btn-sm" onclick="closeBrowse()">Cancel</button>
      <button class="btn btn-sm" id="confirmBtn" onclick="confirmSelect()">Select &amp; Scan</button>
    </div>
  </div>
</div>

<!-- History modal -->
<div class="modal-overlay" id="historyModal">
  <div class="modal history-modal">
    <div class="modal-header">
      <span>Commit History</span>
      <button class="close-btn" onclick="closeHistory()">&times;</button>
    </div>
    <div id="historyAuth" style="padding:12px 16px;border-bottom:1px solid var(--border);display:flex;gap:8px;align-items:center;">
      <span style="color:var(--text-dim);font-size:12px;flex-shrink:0;">SVN Auth:</span>
      <input type="text" id="svnUser" placeholder="username" style="flex:1;background:var(--input-bg);border:1px solid var(--input-border);color:var(--text);padding:4px 8px;border-radius:3px;font-size:12px;outline:none;" />
      <input type="password" id="svnPass" placeholder="password" style="flex:1;background:var(--input-bg);border:1px solid var(--input-border);color:var(--text);padding:4px 8px;border-radius:3px;font-size:12px;outline:none;" />
      <button class="btn btn-sm" onclick="loadLog()">Connect</button>
    </div>
    <div class="history-layout">
      <div class="history-sidebar">
        <div class="log-list" id="logList">
          <div style="padding:30px 16px;color:var(--text-dim);text-align:center;font-size:12px;">Enter SVN credentials above, then click Connect</div>
        </div>
      </div>
      <div class="history-detail">
        <div class="history-detail-header" id="historyDetailHeader">Select a commit to see changed files</div>
        <div class="history-file-list" id="historyFileList"></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn-outline btn btn-sm" onclick="closeHistory()">Close</button>
      <button class="btn btn-sm" id="confirmCommitBtn" onclick="confirmCommit()" style="display:none;">View This Commit</button>
    </div>
  </div>
</div>

<script>
// ---- State ----
let currentPath = '';       // confirmed selected path
let browsePath = '';        // path currently shown in browse modal
let browseData = null;      // last browse API response
let changedFiles = [];
let activeFile = null;
let activeDiffData = null;
let activeSheet = null;

// ---- Utils ----
function esc(s) { const d=document.createElement('div'); d.textContent=String(s); return d.innerHTML; }
function $(sel) { return document.querySelector(sel); }
function $$(sel) { return document.querySelectorAll(sel); }

// ---- localStorage helpers ----
const RECENT_KEY = 'svn_diff_recent_paths';
function getRecentPaths() {
  try { return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); } catch { return []; }
}
function addRecentPath(p) {
  let arr = getRecentPaths().filter(x => x !== p);
  arr.unshift(p);
  if (arr.length > 8) arr = arr.slice(0, 8);
  localStorage.setItem(RECENT_KEY, JSON.stringify(arr));
}

// ---- API ----
async function api(url) {
  const resp = await fetch(url);
  if (!resp.ok) { const d = await resp.json().catch(()=>({})); throw new Error(d.error||`HTTP ${resp.status}`); }
  return resp.json();
}

// ---- Browse modal ----
function openBrowse() {
  $('#browseModal').classList.add('open');
  const initPath = currentPath || '';
  browseTo(initPath);
}
function closeBrowse() { $('#browseModal').classList.remove('open'); }

function confirmSelect() {
  if (!browsePath) return;
  currentPath = browsePath;
  addRecentPath(currentPath);
  const el = $('#pathDisplay');
  el.textContent = currentPath;
  el.classList.remove('empty');
  closeBrowse();
  scanChanges();
}

async function browseTo(path) {
  const body = $('#modalBody');
  body.innerHTML = '<div class="loading"><div class="spinner"></div> Loading...</div>';
  try {
    const data = await api(`/api/browse?path=${encodeURIComponent(path)}`);
    browseData = data;
    browsePath = data.path;
    renderBreadcrumb(data.path, data.sep);
    renderQuickAccess();
    renderDirList(data);
    $('#selectedLabel').textContent = data.path;
  } catch (err) {
    body.innerHTML = `<div style="padding:20px 16px;color:var(--deleted-text);">${esc(err.message)}</div>`;
  }
}

function renderBreadcrumb(fullPath, sep) {
  const bar = $('#breadcrumbBar');
  const parts = fullPath.split(sep).filter(Boolean);
  let html = '';

  // Root
  const isWindows = sep === '\\';
  if (!isWindows) {
    html += `<span class="breadcrumb-item" onclick="browseTo('/')">/</span>`;
  }

  for (let i = 0; i < parts.length; i++) {
    const partial = isWindows
      ? parts.slice(0, i+1).join(sep) + (i === 0 ? sep : '')
      : sep + parts.slice(0, i+1).join(sep);
    const isLast = i === parts.length - 1;
    if (i > 0 || !isWindows) html += `<span class="breadcrumb-sep">&rsaquo;</span>`;
    html += `<span class="breadcrumb-item ${isLast ? 'current' : ''}" ${isLast ? '' : `onclick="browseTo('${esc(partial)}')"`}>${esc(parts[i])}</span>`;
  }
  bar.innerHTML = html;
}

async function renderQuickAccess() {
  const container = $('#quickAccess');
  let html = '';

  // Drives / system roots
  try {
    const data = await api('/api/drives');
    for (const d of data.drives) {
      html += `<button class="quick-btn" onclick="browseTo('${esc(d.path)}')"><span class="qicon">&#128187;</span> ${esc(d.name)}</button>`;
    }
  } catch {}

  // Recent paths
  const recent = getRecentPaths();
  if (recent.length) {
    html += `<span class="quick-section-label">Recent</span>`;
    for (const p of recent) {
      const name = p.split('/').pop() || p.split('\\').pop() || p;
      html += `<button class="quick-btn" onclick="browseTo('${esc(p)}')" title="${esc(p)}"><span class="qicon">&#128338;</span> ${esc(name)}</button>`;
    }
  }
  container.innerHTML = html;
}

function renderDirList(data) {
  const body = $('#modalBody');
  const dirs = data.entries.filter(e => e.is_dir);

  if (!dirs.length) {
    body.innerHTML = '<div style="padding:30px 16px;text-align:center;color:var(--text-dim);">No subdirectories</div>';
    return;
  }

  // Check which subdirs are SVN working copies (async, progressively update)
  let html = '';
  for (const entry of dirs) {
    html += `<div class="dir-item" data-path="${esc(entry.path)}" onclick="browseTo('${esc(entry.path)}')">
      <span class="dir-icon">&#128193;</span>
      <span class="dir-name">${esc(entry.name)}</span>
    </div>`;
  }
  body.innerHTML = html;

  // If current directory is SVN, get status and group changes by subdirectory
  if (data.is_svn) {
    (async () => {
      try {
        const st = await api(`/api/svn/status?path=${encodeURIComponent(data.path)}`);
        if (st.count > 0) {
          const dirCounts = {};
          const sep = data.sep || '/';
          for (const f of st.files) {
            const parts = f.rel_path.split(sep);
            if (parts.length > 1) {
              const subdir = parts[0];
              dirCounts[subdir] = (dirCounts[subdir] || 0) + 1;
            }
          }
          for (const [subdir, count] of Object.entries(dirCounts)) {
            const fullPath = data.path + sep + subdir;
            const el = body.querySelector(`.dir-item[data-path="${CSS.escape(fullPath)}"]`);
            if (el && !el.querySelector('.changes-tag')) {
              el.insertAdjacentHTML('beforeend', `<span class="changes-tag">${count} changes</span>`);
            }
          }
        }
      } catch {}
    })();
  }

  // Always check subdirs for independent SVN checkouts (they may have their own .svn)
  dirs.forEach(async (entry) => {
    try {
      const sub = await api(`/api/browse?path=${encodeURIComponent(entry.path)}`);
      const el = body.querySelector(`.dir-item[data-path="${CSS.escape(entry.path)}"]`);
      if (!el) return;
      // Show SVN tag if this subdir is an independent SVN root (has its own .svn)
      if (sub.is_svn && !el.querySelector('.svn-tag')) {
        el.insertAdjacentHTML('beforeend', '<span class="svn-tag">SVN</span>');
      }
      // Show changes count if this subdir has its own SVN status
      if (sub.is_svn && !el.querySelector('.changes-tag')) {
        try {
          const st = await api(`/api/svn/status?path=${encodeURIComponent(entry.path)}`);
          if (st.count > 0) {
            el.insertAdjacentHTML('beforeend', `<span class="changes-tag">${st.count} changes</span>`);
          }
        } catch {}
      }
    } catch {}
  });
}

// ---- Scan ----
async function scanChanges() {
  if (!currentPath) { openBrowse(); return; }

  $('#fileList').innerHTML = '<div class="loading"><div class="spinner"></div> Scanning...</div>';
  $('#contentArea').innerHTML = '<div class="content-placeholder"><span>Scanning for changes...</span></div>';

  try {
    const data = await api(`/api/svn/status?path=${encodeURIComponent(currentPath)}`);
    changedFiles = data.files;
    $('#fileCount').textContent = data.count;
    $('#refreshBtn').style.display = '';

    if (!data.files.length) {
      $('#fileList').innerHTML = '<div class="content-placeholder" style="padding:30px 14px;font-size:12px;"><span>No changes found</span></div>';
      $('#contentArea').innerHTML = '<div class="content-placeholder"><span>No uncommitted changes in this directory</span></div>';
      return;
    }
    renderFileList();
  } catch (err) {
    $('#fileList').innerHTML = `<div style="padding:20px 14px;color:var(--deleted-text);">${esc(err.message)}</div>`;
  }
}

function refreshChanges() { scanChanges(); }

function renderFileList() {
  const html = changedFiles.map((f, i) => {
    // Show the directory portion of rel_path (e.g. "datas/" from "datas/foo.xls")
    const dirPath = f.rel_path.substring(0, f.rel_path.length - f.name.length);
    return `<div class="file-item" data-index="${i}" onclick="selectFile(${i})">
      <span class="status-badge ${f.status}">${f.status}</span>
      <div class="file-item-info">
        <div class="file-name">${esc(f.name)}</div>
        <div class="file-path">${esc(dirPath || './')}</div>
      </div>
    </div>`;
  }).join('');
  $('#fileList').innerHTML = html;
}

// ---- File selection ----
async function selectFile(index) {
  const f = changedFiles[index];
  activeFile = f;
  $$('.file-item').forEach(el => el.classList.remove('active'));
  $$(`.file-item[data-index="${index}"]`).forEach(el => el.classList.add('active'));
  $('#contentArea').innerHTML = '<div class="loading"><div class="spinner"></div> Computing diff...</div>';

  try {
    const data = await api(`/api/svn/diff?file=${encodeURIComponent(f.path)}`);
    activeDiffData = data;
    if (data.type === 'text') renderTextDiff(data);
    else renderExcelDiff(data);
  } catch (err) {
    $('#contentArea').innerHTML = `<div class="content-placeholder" style="color:var(--deleted-text);">${esc(err.message)}</div>`;
  }
}

// ---- Text diff ----
function renderTextDiff(data) {
  const lines = (data.diff||'').split('\n').map(line => {
    if (line.startsWith('+') && !line.startsWith('+++')) return `<span class="line-add">${esc(line)}</span>`;
    if (line.startsWith('-') && !line.startsWith('---')) return `<span class="line-del">${esc(line)}</span>`;
    if (line.startsWith('@@')) return `<span class="line-hunk">${esc(line)}</span>`;
    return esc(line);
  }).join('\n');
  $('#contentArea').innerHTML = `<div class="text-diff">${lines}</div>`;
}

// ---- Excel diff ----
function renderExcelDiff(data) {
  const sheets = data.sheets;
  const sheetNames = Object.keys(sheets);
  activeSheet = sheetNames[0];
  let html = '<div class="sheet-tabs">';
  for (const name of sheetNames) {
    const sdata = sheets[name];
    const rows = sdata.rows;
    const a = rows.filter(r=>r.type==='added').length;
    const d = rows.filter(r=>r.type==='deleted').length;
    const m = rows.filter(r=>r.type==='modified').length;
    const hasChanges = a+d+m > 0;
    const parts = []; if(a) parts.push(`+${a}`); if(d) parts.push(`-${d}`); if(m) parts.push(`~${m}`);
    const statsText = hasChanges ? parts.join(' ') : 'clean';
    const statsClass = hasChanges ? '' : 'clean';
    const activeClass = name===activeSheet ? 'active' : '';
    html += `<div class="sheet-tab ${activeClass}" data-sheet="${esc(name)}" onclick="switchSheet('${esc(name)}')">${esc(name)}<span class="tab-stats ${statsClass}">${statsText}</span></div>`;
  }
  html += '</div>';
  for (const name of sheetNames) {
    const display = name===activeSheet ? 'flex' : 'none';
    html += `<div class="sheet-panel" data-sheet="${esc(name)}" style="display:${display};flex-direction:column;flex:1;overflow:hidden;">`;
    html += renderSheetDiff(sheets[name], name);
    html += '</div>';
  }
  $('#contentArea').innerHTML = html;
  requestAnimationFrame(syncVisibleDiffRowHeights);
}

function switchSheet(name) {
  activeSheet = name;
  $$('.sheet-tab').forEach(t=>t.classList.toggle('active',t.dataset.sheet===name));
  $$('.sheet-panel').forEach(p=>p.style.display=p.dataset.sheet===name?'flex':'none');
  requestAnimationFrame(syncVisibleDiffRowHeights);
}

let _scrollSyncId = 0;
const CONTEXT_LINES = 3; // rows of context around changes

function renderSheetDiff(sdata, sheetName) {
  const rows=sdata.rows, headers=sdata.headers||[], baseHeaders=sdata.base_headers||[];
  const numCols = Math.max(headers.length, baseHeaders.length);
  const added=rows.filter(r=>r.type==='added').length;
  const deleted=rows.filter(r=>r.type==='deleted').length;
  const modified=rows.filter(r=>r.type==='modified').length;

  const sid = _scrollSyncId++;
  let html = `<div class="filter-bar">
    <span class="filter-label">Filter:</span>
    <button class="filter-btn" data-type="added" onclick="toggleFilter(this)">+ Added (${added})</button>
    <button class="filter-btn" data-type="deleted" onclick="toggleFilter(this)">- Deleted (${deleted})</button>
    <button class="filter-btn" data-type="modified" onclick="toggleFilter(this)">~ Modified (${modified})</button>
    <div class="stats-summary">
      <span class="stat-chip added">+${added}</span>
      <span class="stat-chip deleted">&minus;${deleted}</span>
      <span class="stat-chip modified">~${modified}</span>
    </div>
  </div>`;

  const thBase = Array.from({length:numCols},(_,c)=>`<th>${esc(baseHeaders[c]||'')}</th>`).join('');
  const thWork = Array.from({length:numCols},(_,c)=>`<th>${esc(headers[c]||'')}</th>`).join('');

  // Mark which row indices should be visible (changes + context)
  const visible = new Uint8Array(rows.length);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].type !== 'equal') {
      for (let j = Math.max(0, i - CONTEXT_LINES); j <= Math.min(rows.length - 1, i + CONTEXT_LINES); j++) {
        visible[j] = 1;
      }
    }
  }

  // Build display list: visible rows + collapsed sections
  // Store raw data so expand can work
  const displayChunks = []; // {type:'rows', items:[...]} or {type:'collapsed', count, fromBase, toBase, fromWork, toWork, rowStart, rowEnd}
  let baseLine = 0, workLine = 0;
  let i = 0;
  while (i < rows.length) {
    if (visible[i]) {
      // Emit this row
      const row = rows[i];
      const t = row.type;
      if (t === 'equal') { baseLine++; workLine++; }
      else if (t === 'deleted') { baseLine++; }
      else if (t === 'added') { workLine++; }
      else { baseLine++; workLine++; }
      displayChunks.push({ type: 'row', row, baseLine, workLine, idx: i });
      i++;
    } else {
      // Collapse consecutive hidden (equal) rows
      const fromBase = baseLine + 1;
      const fromWork = workLine + 1;
      const startIdx = i;
      while (i < rows.length && !visible[i]) {
        baseLine++; workLine++; // all hidden rows are 'equal'
        i++;
      }
      const count = i - startIdx;
      displayChunks.push({ type: 'collapsed', count, fromBase, toBase: baseLine, fromWork, toWork: workLine, rowStart: startIdx, rowEnd: i - 1 });
    }
  }

  // Render chunks
  let leftRows = '', rightRows = '';
  for (const chunk of displayChunks) {
    if (chunk.type === 'collapsed') {
      const collapseId = `exp_${sid}_${chunk.rowStart}`;
      const pairId = `${sid}_c${chunk.rowStart}`;
      leftRows  += `<tr class="row-collapsed" data-expand="${collapseId}" data-pair="${pairId}"><td colspan="${numCols+2}" class="collapse-cell" onclick="expandRows('${collapseId}', ${sid})">&#8943; ${chunk.count} unchanged rows (${chunk.fromBase}-${chunk.toBase}) &#8943;</td></tr>`;
      rightRows += `<tr class="row-collapsed" data-expand="${collapseId}" data-pair="${pairId}"><td colspan="${numCols+2}" class="collapse-cell" onclick="expandRows('${collapseId}', ${sid})">&#8943; ${chunk.count} unchanged rows (${chunk.fromWork}-${chunk.toWork}) &#8943;</td></tr>`;
    } else {
      const row = chunk.row;
      const t = row.type;
      const changed = new Set(row.changed_cols || []);
      const bl = chunk.baseLine, wl = chunk.workLine;
      const pairId = `${sid}_r${chunk.idx}`;
      if (t === 'equal') {
        leftRows  += `<tr class="row-equal" data-type="equal" data-pair="${pairId}"><td class="ln">${bl}</td><td class="gutter"></td>${cellsHtml(row.base,numCols)}</tr>`;
        rightRows += `<tr class="row-equal" data-type="equal" data-pair="${pairId}"><td class="ln">${wl}</td><td class="gutter"></td>${cellsHtml(row.work,numCols)}</tr>`;
      } else if (t === 'deleted') {
        leftRows  += `<tr class="row-deleted" data-type="deleted" data-pair="${pairId}"><td class="ln">${bl}</td><td class="gutter"></td>${cellsHtml(row.base,numCols)}</tr>`;
        rightRows += `<tr class="row-deleted" data-type="deleted" data-pair="${pairId}"><td class="ln empty-side"></td><td class="gutter"></td>${emptyCells(numCols)}</tr>`;
      } else if (t === 'added') {
        leftRows  += `<tr class="row-added" data-type="added" data-pair="${pairId}"><td class="ln empty-side"></td><td class="gutter"></td>${emptyCells(numCols)}</tr>`;
        rightRows += `<tr class="row-added" data-type="added" data-pair="${pairId}"><td class="ln">${wl}</td><td class="gutter"></td>${cellsHtml(row.work,numCols)}</tr>`;
      } else if (t === 'modified') {
        let lc='',rc='';
        for(let c=0;c<numCols;c++){
          const bv=row.base&&c<row.base.length?row.base[c]:'';
          const wv=row.work&&c<row.work.length?row.work[c]:'';
          if(changed.has(c)){lc+=`<td class="cell-changed"><span class="old-val">${escVal(bv)}</span></td>`;rc+=`<td class="cell-changed"><span class="new-val">${escVal(wv)}</span></td>`;}
          else{lc+=`<td>${escVal(bv)}</td>`;rc+=`<td>${escVal(wv)}</td>`;}
        }
        leftRows  += `<tr class="row-modified" data-type="modified" data-pair="${pairId}"><td class="ln">${bl}</td><td class="gutter"></td>${lc}</tr>`;
        rightRows += `<tr class="row-modified" data-type="modified" data-pair="${pairId}"><td class="ln">${wl}</td><td class="gutter"></td>${rc}</tr>`;
      }
    }
  }

  // Store row data globally for expand
  window[`_diffData_${sid}`] = { rows, numCols };

  // Labels row (fixed, not scrolling)
  html += `<div class="diff-labels">`;
  html += `<div class="diff-label left">BASE (SVN)</div>`;
  html += `<div class="diff-label">WORKING COPY</div>`;
  html += `</div>`;
  // Shared vertical scroll container with two horizontal-scroll halves
  html += `<div class="diff-scroll" id="diffScroll${sid}">`;
  html += `<div class="diff-half left" id="diffL${sid}"><table class="diff-table">`;
  html += `<thead><tr><th>#</th><th></th>${thBase}</tr></thead>`;
  html += `<tbody>${leftRows}</tbody></table></div>`;
  html += `<div class="diff-half right" id="diffR${sid}"><table class="diff-table">`;
  html += `<thead><tr><th>#</th><th></th>${thWork}</tr></thead>`;
  html += `<tbody>${rightRows}</tbody></table></div>`;
  html += `</div>`;

  // Sync horizontal scroll only (vertical is shared via .diff-scroll)
  requestAnimationFrame(() => {
    setupHScrollSync(`diffL${sid}`, `diffR${sid}`);
    syncDiffRowHeights(sid);
  });
  return html;
}

function expandRows(collapseId, sid) {
  const dd = window[`_diffData_${sid}`];
  if (!dd) return;
  const leftTable = $(`#diffL${sid} tbody`);
  const rightTable = $(`#diffR${sid} tbody`);
  const leftCollapsed = leftTable.querySelector(`tr[data-expand="${collapseId}"]`);
  const rightCollapsed = rightTable.querySelector(`tr[data-expand="${collapseId}"]`);
  if (!leftCollapsed || !rightCollapsed) return;

  // Parse rowStart from collapseId: "exp_{sid}_{rowStart}"
  const rowStart = parseInt(collapseId.split('_').pop());
  const rows = dd.rows, numCols = dd.numCols;

  // Find consecutive equal rows from rowStart
  let newLeftHtml = '', newRightHtml = '';
  let bl = 0, wl = 0;
  // Count lines up to rowStart
  for (let i = 0; i < rowStart; i++) {
    const t = rows[i].type;
    if (t === 'equal' || t === 'modified') { bl++; wl++; }
    else if (t === 'deleted') { bl++; }
    else if (t === 'added') { wl++; }
  }
  let i = rowStart;
  while (i < rows.length && rows[i].type === 'equal') {
    bl++; wl++;
    const pairId = `${sid}_r${i}`;
    newLeftHtml  += `<tr class="row-equal" data-type="equal" data-pair="${pairId}"><td class="ln">${bl}</td><td class="gutter"></td>${cellsHtml(rows[i].base, numCols)}</tr>`;
    newRightHtml += `<tr class="row-equal" data-type="equal" data-pair="${pairId}"><td class="ln">${wl}</td><td class="gutter"></td>${cellsHtml(rows[i].work, numCols)}</tr>`;
    i++;
  }

  // Replace collapsed row with expanded rows
  leftCollapsed.insertAdjacentHTML('afterend', newLeftHtml);
  leftCollapsed.remove();
  rightCollapsed.insertAdjacentHTML('afterend', newRightHtml);
  rightCollapsed.remove();

  // Re-sync horizontal scroll
  setupHScrollSync(`diffL${sid}`, `diffR${sid}`);
  requestAnimationFrame(() => syncDiffRowHeights(sid));
}

function syncVisibleDiffRowHeights() {
  $$('.sheet-panel').forEach(panel => {
    if (panel.style.display === 'none') return;
    panel.querySelectorAll('.diff-scroll').forEach(scroll => {
      const sid = scroll.id.replace('diffScroll', '');
      syncDiffRowHeights(sid);
    });
  });
}

function syncDiffRowHeights(sid) {
  const leftBody = document.querySelector(`#diffL${sid} tbody`);
  const rightBody = document.querySelector(`#diffR${sid} tbody`);
  if (!leftBody || !rightBody || !leftBody.offsetParent) return;

  leftBody.querySelectorAll('tr[data-pair]').forEach(leftRow => {
    const pair = leftRow.dataset.pair;
    const rightRow = rightBody.querySelector(`tr[data-pair="${pair}"]`);
    if (!rightRow) return;

    leftRow.style.height = '';
    rightRow.style.height = '';
    leftRow.style.minHeight = '';
    rightRow.style.minHeight = '';

    if (leftRow.style.display === 'none' || rightRow.style.display === 'none') return;
    const height = Math.max(leftRow.getBoundingClientRect().height, rightRow.getBoundingClientRect().height);
    if (height > 0) {
      const px = `${Math.ceil(height)}px`;
      leftRow.style.height = px;
      rightRow.style.height = px;
      leftRow.style.minHeight = px;
      rightRow.style.minHeight = px;
    }
  });
}

function setupHScrollSync(leftId, rightId) {
  const left = document.getElementById(leftId);
  const right = document.getElementById(rightId);
  if (!left || !right) return;
  // Only sync horizontal scroll — vertical is shared by parent .diff-scroll
  let syncing = false;
  left.addEventListener('scroll', () => {
    if (syncing) return;
    syncing = true;
    right.scrollLeft = left.scrollLeft;
    requestAnimationFrame(() => { syncing = false; });
  });
  right.addEventListener('scroll', () => {
    if (syncing) return;
    syncing = true;
    left.scrollLeft = right.scrollLeft;
    requestAnimationFrame(() => { syncing = false; });
  });
}

function escVal(v) { if(v===''||v===null||v===undefined) return '<span style="opacity:0.3">-</span>'; return esc(v); }
function cellsHtml(data,n) { if(!data) return emptyCells(n); return Array.from({length:n},(_,c)=>`<td>${escVal(c<data.length?data[c]:'')}</td>`).join(''); }
function emptyCells(n) { return Array.from({length:n},()=>'<td class="empty-side">&nbsp;</td>').join(''); }

function toggleFilter(btn) {
  btn.classList.toggle('active');
  const panel = btn.closest('.sheet-panel') || btn.closest('.content');
  const buttons = panel.querySelectorAll('.filter-btn');
  const active = {};
  buttons.forEach(b => active[b.dataset.type] = b.classList.contains('active'));
  const anyActive = Object.values(active).some(v => v);
  // Apply filter to both left and right tables simultaneously
  const tables = panel.querySelectorAll('.diff-table');
  tables.forEach(table => {
    table.querySelectorAll('tbody tr').forEach(row => {
      if (!anyActive) { row.style.display=''; return; }
      row.style.display = active[row.dataset.type] ? '' : 'none';
    });
  });
  requestAnimationFrame(syncVisibleDiffRowHeights);
}

// ---- History ----
let historyWcPath = '';

function svnAuth() {
  const u = $('#svnUser').value.trim();
  const p = $('#svnPass').value.trim();
  if (u && p) return `&username=${encodeURIComponent(u)}&password=${encodeURIComponent(p)}`;
  return '';
}

function openHistory() {
  if (!currentPath) { openBrowse(); return; }
  historyWcPath = currentPath;
  $('#historyModal').classList.add('open');
  $('#historyDetailHeader').textContent = 'Select a commit to see changed files';
  $('#historyFileList').innerHTML = '';
  // Auto-load if credentials already filled
  if ($('#svnUser').value && $('#svnPass').value) loadLog();
}
function closeHistory() { $('#historyModal').classList.remove('open'); }

// Allow Enter key in auth fields to trigger connect
document.addEventListener('DOMContentLoaded', () => {
  ['svnUser','svnPass'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('keydown', e => { if (e.key === 'Enter') loadLog(); });
  });
});

async function loadLog() {
  const list = $('#logList');
  list.innerHTML = '<div class="loading"><div class="spinner"></div> Loading history...</div>';
  try {
    const data = await api(`/api/svn/log?path=${encodeURIComponent(historyWcPath)}&limit=50${svnAuth()}`);
    if (!data.entries.length) {
      list.innerHTML = '<div style="padding:20px 16px;color:var(--text-dim);">No history found</div>';
      return;
    }
    list.innerHTML = data.entries.map(e => {
      const date = e.date ? e.date.substring(0,10) + ' ' + e.date.substring(11,16) : '';
      const msg = e.message ? e.message.split('\n')[0] : '(no message)';
      return `<div class="log-item" data-rev="${esc(e.revision)}" onclick="selectCommit('${esc(e.revision)}', this)">
        <div><span class="log-rev">r${esc(e.revision)}</span><span class="log-meta">${esc(e.author)} &middot; ${esc(date)}</span></div>
        <div class="log-msg">${esc(msg)}</div>
      </div>`;
    }).join('');
  } catch (err) {
    list.innerHTML = `<div style="padding:20px 16px;color:var(--deleted-text);">${esc(err.message)}</div>`;
  }
}

let selectedCommitRev = null;
let selectedCommitFiles = [];

async function selectCommit(rev, el) {
  $$('.log-item').forEach(e => e.classList.remove('active'));
  el.classList.add('active');
  selectedCommitRev = null;
  selectedCommitFiles = [];
  $('#confirmCommitBtn').style.display = 'none';

  $('#historyDetailHeader').innerHTML = `<div class="loading" style="padding:4px;"><div class="spinner"></div> Loading r${esc(rev)}...</div>`;
  $('#historyFileList').innerHTML = '';

  try {
    const data = await api(`/api/svn/log-detail?path=${encodeURIComponent(historyWcPath)}&revision=${rev}${svnAuth()}`);
    const date = data.date ? data.date.substring(0,10) + ' ' + data.date.substring(11,16) : '';
    $('#historyDetailHeader').innerHTML = `<strong>r${esc(rev)}</strong> &middot; ${esc(data.author)} &middot; ${esc(date)}<br><span style="color:var(--text);">${esc(data.message || '(no message)')}</span>`;

    if (!data.files.length) {
      $('#historyFileList').innerHTML = '<div style="padding:20px;color:var(--text-dim);">No files changed</div>';
      return;
    }

    selectedCommitRev = rev;
    selectedCommitFiles = data.files;
    $('#confirmCommitBtn').style.display = '';

    // Preview file list (read-only, no onclick)
    $('#historyFileList').innerHTML = data.files.map(f => `
      <div class="history-file-item" style="cursor:default;">
        <span class="action-badge ${f.action}">${f.action}</span>
        <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(f.name)}</span>
        <span style="color:var(--text-dim);font-size:11px;margin-left:auto;flex-shrink:0;">${esc(f.path)}</span>
      </div>
    `).join('');
  } catch (err) {
    $('#historyDetailHeader').textContent = err.message;
  }
}

function confirmCommit() {
  if (!selectedCommitRev || !selectedCommitFiles.length) return;
  const rev = selectedCommitRev;
  const files = selectedCommitFiles;
  closeHistory();

  // Map action letters to status-badge compatible format
  const actionToStatus = { A: 'A', D: 'D', M: 'M', R: 'M' };

  // Populate sidebar with this commit's files
  changedFiles = files.map(f => ({
    status: actionToStatus[f.action] || 'M',
    path: f.path,
    name: f.name,
    rel_path: f.path,
    type: /\.(xls|xlsx)$/i.test(f.name) ? 'excel' : 'text',
    _historyRev: rev,
  }));
  $('#fileCount').textContent = changedFiles.length;
  $('#pathDisplay').textContent = `${currentPath}  [r${rev}]`;
  $('#pathDisplay').classList.remove('empty');

  // Render file list in sidebar
  const html = changedFiles.map((f, i) => {
    const dirPath = f.path.substring(0, f.path.length - f.name.length);
    return `<div class="file-item" data-index="${i}" onclick="selectHistoryFile(${i})">
      <span class="status-badge ${f.status}">${f.status}</span>
      <div class="file-item-info">
        <div class="file-name">${esc(f.name)}</div>
        <div class="file-path">${esc(dirPath)}</div>
      </div>
    </div>`;
  }).join('');
  $('#fileList').innerHTML = html;

  // Reset content
  $('#contentArea').innerHTML = '<div class="content-placeholder"><span>Select a file to view its diff at r' + esc(rev) + '</span></div>';
}

async function selectHistoryFile(index) {
  const f = changedFiles[index];
  $$('.file-item').forEach(el => el.classList.remove('active'));
  $$(`.file-item[data-index="${index}"]`).forEach(el => el.classList.add('active'));
  const rev = f._historyRev;

  $('#contentArea').innerHTML = '<div class="loading"><div class="spinner"></div> Loading diff for r' + esc(rev) + '...</div>';

  try {
    const data = await api(`/api/svn/history-diff?path=${encodeURIComponent(historyWcPath)}&file=${encodeURIComponent(f.path)}&revision=${rev}${svnAuth()}`);
    if (data.type === 'text') {
      renderTextDiff(data);
    } else {
      renderExcelDiff(data);
    }
  } catch (err) {
    $('#contentArea').innerHTML = `<div class="content-placeholder" style="color:var(--deleted-text);">${esc(err.message)}</div>`;
  }
}
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
