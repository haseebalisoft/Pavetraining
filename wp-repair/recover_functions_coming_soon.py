"""Recover pave-training functions.php: strip broken Coming Soon, append clean block."""
import paramiko
import sys
from pathlib import Path

host = "1236623.eu11.ssh.myftpupload.com"
password = "axSMWYGJVAifcLfvn9X5spSU"
users = ["wpusername7000", "1236623", "u1236623", "pave"]
ports = [22, 2222]
remote_candidates = [
    "wp-content/themes/pave-training/functions.php",
    "html/wp-content/themes/pave-training/functions.php",
    "public_html/wp-content/themes/pave-training/functions.php",
]
block_path = Path(__file__).with_name("coming-soon-block.php")
local_tmp = Path(__file__).with_name("functions.php.recovered")
block = block_path.read_text(encoding="utf-8")

for port in ports:
    for user in users:
        try:
            t = paramiko.Transport((host, port))
            t.banner_timeout = 60
            t.connect(username=user, password=password)
            sftp = paramiko.SFTPClient.from_transport(t)
            print(f"AUTH OK user={user} port={port}")
            print("cwd:", sftp.listdir(".")[:30])
            for remote in remote_candidates:
                try:
                    with sftp.open(remote, "r") as f:
                        data = f.read().decode("utf-8", errors="replace")
                    print(f"READ {remote} bytes={len(data)}")
                    marker = "/* PAVE COMING SOON"
                    p = data.find(marker)
                    if p < 0:
                        p = data.find("PAVE COMING SOON")
                    if p >= 0:
                        data = data[:p].rstrip() + "\n\n"
                        print(f"Truncated at {p}")
                    else:
                        data = data.rstrip() + "\n\n"
                        print("No marker; appending")
                    data = data + block
                    if not data.startswith("<?php"):
                        raise SystemExit("Refusing to write: missing PHP open tag")
                    local_tmp.write_text(data, encoding="utf-8")
                    sftp.put(str(local_tmp), remote)
                    print("UPLOADED", remote, "bytes", len(data))
                    sftp.close()
                    t.close()
                    sys.exit(0)
                except FileNotFoundError:
                    print("missing", remote)
                except Exception as e:
                    print("try fail", remote, type(e).__name__, e)
            sftp.close()
            t.close()
        except SystemExit:
            raise
        except Exception as e:
            print(f"AUTH FAIL user={user} port={port}: {type(e).__name__}: {e}")

print("ALL FAILED")
sys.exit(1)
