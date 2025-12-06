# docx_redline Development Plan

## Project Overview

A high-level Python API for editing Word documents with tracked changes, eliminating the need to write raw OOXML XML.

**Goal**: Reduce code from 30+ lines of XML manipulation to 3 lines of high-level API calls.

**Success Criteria**: Complete all 11 surgical edits from `examples/surgical_edits.yaml` without writing XML.

## Project Structure

```
~/Projects/docx_redline/
├── docs/                        # Complete API specification (162 pages)
│   ├── PROPOSED_API_README.md  # Executive summary
│   ├── PROPOSED_API.md         # Full API spec (92 pages)
│   ├── IMPLEMENTATION_NOTES.md # Technical guide (45 pages)
│   └── QUICK_REFERENCE.md      # Quick lookup (15 pages)
├── examples/                    # YAML edit specifications
│   ├── surgical_edits.yaml     # Real-world 11-edit scenario
│   ├── simple_edits.yaml       # Basic contract edits
│   └── citation_updates.yaml   # Batch citation updates
├── src/
│   └── docx_redline/           # Main package (to be built)
├── tests/                       # Test suite (to be built)
└── .beads/                      # Task tracking database
```

## Epic 1: MVP - Core Text Operations

### Phase Overview

Build the core functionality to enable surgical document edits without writing XML. This epic focuses on text operations (insert, replace, delete) with basic scope support.

### Critical Path

1. **Setup** (docx_redline-vvi)
   - Package structure with pyproject.toml
   - Dependencies: lxml, python-dateutil, pyyaml
   - Modern Python packaging (Python 3.10+)

2. **Foundation** (docx_redline-uj2)
   - Custom exception classes with helpful error messages
   - Base for all error handling

3. **Core Algorithm #1** (docx_redline-ye8) - **HIGH PRIORITY**
   - Text search with fragmentation handling
   - This is THE critical piece that makes everything work
   - Algorithm: build character map across runs, find text, map back

4. **Core Algorithm #2** (docx_redline-03e) - **HIGH PRIORITY**
   - TrackedXMLGenerator for insertions
   - Auto-generate <w:ins> XML with proper attributes
   - Handle timestamps, IDs, RSID, xml:space

5. **Integration** (docx_redline-c3m)
   - Minimal Document class
   - Get ONE edit working end-to-end
   - Prove the concept

6. **Expansion**
   - TextSpan class (docx_redline-yk0)
   - Scope system (docx_redline-klh)
   - Replace/delete operations (docx_redline-b23)
   - Error suggestions (docx_redline-wq7)

7. **Batch Processing**
   - apply_edits() (docx_redline-3dx)
   - YAML support (docx_redline-lvb)

8. **Utilities**
   - accept_all_changes() (docx_redline-vtz)
   - delete_all_comments()

9. **Validation** (docx_redline-47a) - **HIGH PRIORITY**
   - Integration test with all 11 surgical edits
   - This is the success criteria

### Development Strategy: Fastest Path

**Week 1: Proof of Concept**
- Day 1-2: Setup + Text search algorithm
- Day 3-4: XML generator + minimal Document class
- Day 5: Get ONE edit working (insert_tracked)

**Week 2: Complete Operations**
- Day 1-2: TextSpan, Scope system
- Day 3-4: Replace/delete operations
- Day 5: Batch processing

**Week 3: Polish & Validate**
- Day 1-2: Error handling, utilities
- Day 3-4: YAML support, testing
- Day 5: Integration test with all 11 edits

## Current Status

✅ Project structure created
✅ Documentation copied (162 pages)
✅ Examples copied (3 YAML files)
✅ Beads task tracker initialized
✅ 16 issues created and prioritized

**Ready to work on**: Setup, Foundation, Core algorithms

## Next Actions

1. Claim issue: `bd claim docx_redline-vvi` (Setup)
2. Create package structure
3. Move to core algorithms

## Resources

- **API Spec**: `docs/PROPOSED_API.md`
- **Implementation Guide**: `docs/IMPLEMENTATION_NOTES.md`
- **Quick Reference**: `docs/QUICK_REFERENCE.md`
- **Real Example**: `examples/surgical_edits.yaml`

## Track Progress

```bash
cd ~/Projects/docx_redline

# See all issues
bd list

# See ready work
bd ready

# Claim an issue
bd claim <issue-id>

# Mark complete
bd close <issue-id>
```
