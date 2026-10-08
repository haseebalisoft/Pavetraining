"""
Upload Coming Soon fix via SFTP.
Set env vars, then run:
  set SFTP_HOST=...
  set SFTP_USER=...
  set SFTP_PASS=...
  set SFTP_PORT=22
  python wp-repair/upload_clean_functions.py

Or pass as args:
  python wp-repair/upload_clean_functions.py HOST USER PASS [PORT]
"""
import os
import sys
from pathlib import Path

import paramiko

ROOT = Path(__file__).resolve().parent
BLOCK = (ROOT / "coming-soon-block.php").read_text(encoding="utf-8")
OUT = ROOT / "functions.php.recovered"

REMOTE_CANDIDATES = [
    "wp-content/themes/pave-training/functions.php",
    "html/wp-content/themes/pave-training/functions.php",
    "public_html/wp-content/themes/pave-training/functions.php",
    "sites/pavetraining.co.uk/wp-content/themes/pave-training/functions.php",
]


def recover(data: str) -> str:
    for marker in ("/* PAVE COMING SOON", "PAVE COMING SOON", "PAVE_PREVIEW_SECREV", "define('PAVE_PREVIEW_SECRET'"):
        p = data.find(marker)
        if p >= 0:
            # If marker is mid-define from corruption, walk back to comment or previous newline block
            cut = data.rfind("\n/* PAVE COMING SOON", 0, p + 1)
            if cut < 0:
                cut = data.rfind("\n/* PAVE COMING SOON REMOVED", 0, p + 1)
            if cut < 0:
                cut = p
            data = data[:cut].rstrip() + "\n\n"
            break
    if not data.lstrip().startswith("<?php"):
        raise SystemExit("Refusing write: functions.php missing <?php")
    return data.rstrip() + "\n\n" + BLOCK


def main():
    if len(sys.argv) >= 4:
        host, user, password = sys.argv[1], sys.argv[2], sys.argv[3]
        port = int(sys.argv[4]) if len(sys.argv) > 4 else 22
    else:
        host = os.environ.get("SFTP_HOST", "").strip()
        user = os.environ.get("SFTP_USER", "").strip()
        password = os.environ.get("SFTP_PASS", "").strip()
        port = int(os.environ.get("SFTP_PORT", "22"))
    if not host or not user or not password:
        print("Need SFTP_HOST SFTP_USER SFTP_PASS (and optional SFTP_PORT)")
        print("Or: python upload_clean_functions.py HOST USER PASS [PORT]")
        sys.exit(2)

    t = paramiko.Transport((host, port))
    t.banner_timeout = 90
    t.connect(username=user, password=password)
    sftp = paramiko.SFTPClient.from_transport(t)
    print("AUTH OK", host, user, port)
    try:
        print("root listing:", sftp.listdir(".")[:40])
    except Exception as e:
        print("list fail", e)

    uploaded = False
    for remote in REMOTE_CANDIDATES:
        try:
            with sftp.open(remote, "r") as f:
                raw = f.read().decode("utf-8", errors="replace")
            fixed = recover(raw)
            OUT.write_text(fixed, encoding="utf-8")
            sftp.put(str(OUT), remote)
            print("UPLOADED", remote, "bytes", len(fixed))
            uploaded = True
            break
        except FileNotFoundError:
            print("missing", remote)
        except Exception as e:
            print("fail", remote, type(e).__name__, e)

    if not uploaded:
        # walk for pave-training
        def walk(p, depth=0):
            if depth > 4:
                return
            try:
                for name in sftp.listdir(p):
                    full = f"{p}/{name}" if p != "." else name
                    if name in ("wp-content", "themes", "pave-training", "html", "public_html", "sites"):
                        print("DIR", full)
                        walk(full, depth + 1)
                    if name == "functions.php" and "pave-training" in full.replace("\\", "/"):
                        print("FOUND", full)
            except Exception:
                pass

        walk(".")
        sftp.close()
        t.close()
        sys.exit(1)

    sftp.close()
    t.close()
    print("DONE")


if __name__ == "__main__":
    main()
