# Contributing

## Scope and Platform

- Primary runtime is Windows + AutoCAD + Python 3.10+.
- The Python server is in `src/acad_cmd`.
- The optional in-process bridge plugin is in `plugins/AcadEventBridge`.

## Local Setup

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install -U pip
pip install -e .[dev]
```

If `.[dev]` is not available in your environment, use `pip install -e .` and run tests with `unittest`.

## Tests

Run unit tests:

```bat
python -m unittest discover -s tests -p "test_*.py"
```

Integration tests (require a running AutoCAD):

```bat
set ACAD_MCP_RUN_INTEGRATION=1
python -m unittest discover -s tests -p "test_*.py"
```

Smoke scripts are in `scripts/`.

## Pull Requests

- Keep changes focused and small where possible.
- Add or update tests for behavior changes.
- Update docs (`README.md` / `docs/*`) when behavior or configuration changes.
- Do not commit local runtime artifacts (`logs/`, `out/`, `tmp/`, `.venv/`).
- For AutoCAD-specific behavior, include a short manual validation note in the PR.

## Commit Style

- Use clear, imperative commit messages.
- Mention the affected area (`tools`, `bridge`, `docs`, `tests`) in the subject when possible.
