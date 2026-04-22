# Release Notes Template

## `vX.Y.Z` - YYYY-MM-DD

## Summary

Short overview of what this release changes.

## Added

- ...

## Changed

- ...

## Fixed

- ...

## Breaking Changes

- None.

## Upgrade Notes

- Environment variables to review:
  - `AUTOCAD_MCP_TARGET_MAJOR`
  - `AUTOCAD_MCP_EVENT_BRIDGE_ENABLED`
  - `AUTOCAD_MCP_EVENT_BRIDGE_OBJECT_EVENTS_ENABLED`

## Validation

- Unit tests:
  - `python -m unittest discover -s tests -p "test_*.py"`
- Smoke checks:
  - `scripts/mcp_smoketest_stdio.py`
  - `scripts/mcp_baseline_capture.py`
