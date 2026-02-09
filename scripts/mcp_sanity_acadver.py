import os
import time

from acad_cmd.autocad_bridge import AutoCADBridge


def main() -> int:
    b = AutoCADBridge()
    if not b.connect():
        print("connect False")
        return 2

    logp = str(b.get_variable("LOGFILENAME"))
    print(f"LOGFILENAME={logp}")

    # Ensure logging enabled (we use the existing log file path)
    try:
        b.set_variable("LOGFILEMODE", 1)
    except Exception as e:
        print(f"warn: failed to set LOGFILEMODE: {e}")

    try:
        pre = os.path.getsize(logp) if logp and os.path.exists(logp) else 0
    except Exception:
        pre = 0

    # Send a LISP expression that does not require double quotes.
    # getvar accepts a symbol as well as a string.
    expr = "(getvar 'ACADVER)"
    b.send_command(expr)

    # Wait for the log to grow
    for _ in range(100):
        try:
            if os.path.getsize(logp) > pre:
                break
        except Exception:
            pass
        time.sleep(0.1)

    with open(logp, "rb") as f:
        f.seek(pre)
        data = f.read(8192)

    # Best-effort decode
    txt = None
    for enc in ("cp1251", "mbcs", "utf-8"):
        try:
            txt = data.decode(enc)
            break
        except Exception:
            pass
    if txt is None:
        txt = data.decode("latin1", "replace")

    print("---new log---")
    print(txt.strip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
