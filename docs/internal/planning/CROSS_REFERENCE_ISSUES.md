# Cross-Reference Implementation Issues

Created: 2025-12-31

This file tracks the beads issues created for implementing cross-reference support.

## Issue IDs

| Phase | Issue ID | Title | Status |
|-------|----------|-------|--------|
| 1 | docx_redline-898 | Cross-refs Phase 1: Core Infrastructure and Data Models | open |
| 2 | docx_redline-899 | Cross-refs Phase 2: Bookmark Management (Create/Read) | open |
| 3 | docx_redline-900 | Cross-refs Phase 3: Basic Cross-Reference Insertion (Bookmark Targets) | open |
| 4 | docx_redline-901 | Cross-refs Phase 4: Heading References | open |
| 5 | docx_redline-902 | Cross-refs Phase 5: Caption References (Figures and Tables) | open |
| 6 | docx_redline-903 | Cross-refs Phase 6: Note References and Convenience Methods | open |
| 7 | docx_redline-904 | Cross-refs Phase 7: Inspection and Field Management | open |

## Dependency Chain

```
Phase 1 (docx_redline-898)
    |
    +---> Phase 2 (docx_redline-899)
    |         |
    +---------+---> Phase 3 (docx_redline-900)
                        |
                        +---> Phase 4 (docx_redline-901)
                                  |
                                  +---> Phase 5 (docx_redline-902)
                                            |
                                            +---> Phase 6 (docx_redline-903)
                                                      |
                                                      +---> Phase 7 (docx_redline-904)
```

## Quick Reference

```bash
# View issue details
bd --no-db show docx_redline-898  # Phase 1
bd --no-db show docx_redline-899  # Phase 2
bd --no-db show docx_redline-900  # Phase 3
bd --no-db show docx_redline-901  # Phase 4
bd --no-db show docx_redline-902  # Phase 5
bd --no-db show docx_redline-903  # Phase 6
bd --no-db show docx_redline-904  # Phase 7

# Claim a phase
bd --no-db update docx_redline-898 -s in_progress

# Close a phase
bd --no-db close docx_redline-898 -r "completed"

# View dependency tree
bd --no-db dep tree docx_redline-904
```

## Reference

See `CROSS_REFERENCE_DEV_PLAN.md` in this directory for full implementation details.
