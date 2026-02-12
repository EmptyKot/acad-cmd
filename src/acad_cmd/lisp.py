import os


def normalize_path_for_autocad(path: str) -> str:
    # AutoCAD LISP typically accepts forward slashes; also helps escaping.
    p = os.path.abspath(path)
    return p.replace("\\", "/")


def lisp_quote_string(s: str) -> str:
    # Escape backslashes and quotes for an AutoLISP string literal.
    return s.replace("\\", "\\\\").replace('"', '\\"')


def build_load_lisp_command(path: str) -> str:
    p = normalize_path_for_autocad(path)
    return f'(load "{lisp_quote_string(p)}")'


def build_run_lisp_script(expr: str, marker_id: str) -> str:
    # Send multiple LISP lines in one SendCommand call.
    # Markers appear in command history/logfile.
    start = f'(prompt "\\n[MCP:LISP id={marker_id} start]")'
    end = f'(prompt "\\n[MCP:LISP id={marker_id} end]")'
    # Ensure prompt lines are printed.
    return "\n".join([
        start,
        "(princ)",
        expr,
        end,
        "(princ)",
    ])
