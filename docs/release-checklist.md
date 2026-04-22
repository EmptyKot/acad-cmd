# Release Checklist

Use this checklist before creating a public release tag.

## 1. Repository Hygiene

- [ ] Working tree is clean (`git status`).
- [ ] No local/runtime artifacts are tracked (`logs/`, `out/`, `tmp/`, `.venv/`).
- [ ] `README.md`, `CONTRIBUTING.md`, `SECURITY.md`, and `LICENSE` are present and up to date.

## 2. Validation

- [ ] Unit tests pass:
  - `python -m unittest discover -s tests -p "test_*.py"`
- [ ] AutoCAD integration checks run in a real environment (when preparing a production release).
- [ ] Smoke scripts relevant to the release scope have been executed from `scripts/`.

## 3. Versioning

- [ ] `pyproject.toml` version is bumped.
- [ ] Changelog/release notes are prepared (see `docs/release-notes-template.md`).
- [ ] Breaking changes are clearly called out.

## 4. GitHub Release

- [ ] Create and push tag: `vX.Y.Z`.
- [ ] Open GitHub release with release notes.
- [ ] Attach known limitations and migration notes.

## 5. Post-release

- [ ] Verify CI status for tag/`main`.
- [ ] Confirm installation path still works from clean clone:
  - `py -3.11 -m venv .venv`
  - `.venv\Scripts\activate`
  - `pip install .`
