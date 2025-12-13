# Minimal Editing (Legal-Style Diffs) — Product Spec & Implementation Plan

## Summary

Add a “minimal editing” mode for tracked changes generation so that small edits produce small, human-looking redlines (legal-review style), instead of “delete entire old text + insert entire new text”.

The primary target is `Document.compare_to()` (redlining from original vs modified). A follow-on phase optionally extends the same behavior to `replace_tracked()`.

## Problem

Current behavior for paragraph text changes is typically:

- delete the old text (often all of it), then
- insert the new text (often all of it).

This creates noisy redlines that obscure what actually changed, especially in legal documents where edits are usually a few words or punctuation.

## Goals

### G1 — Legal-style tracked diffs

Within a paragraph, generate tracked changes that:

- preserve unchanged text as plain, unchanged runs (outside tracked markup),
- represent replacements as **deletion first, then insertion** at the same location,
- avoid hyper-fragmented character-by-character markup,
- avoid long “delete everything / insert everything” changes when only a few words changed.

### G2 — Maintain “accept changes” correctness (text)

For non-whitespace text changes, accepting tracked changes should produce the modified text for that paragraph.

Explicitly not required:
- If a paragraph differs only by whitespace, minimal mode may emit no changes and accepting changes may not reproduce the modified paragraph’s whitespace exactly (reviewer-friendly behavior).

### G3 — Predictable output and safe fallback

If the diff would be too fragmented or the paragraph is unsafe to edit minimally, fall back to the existing “coarse” behavior for that paragraph (so we do not generate unreadable or risky OOXML).

## Non-Goals (MVP)

- Tracking formatting-only changes (bold/italic/style changes) as tracked formatting revisions.
- Structural diffs beyond paragraph alignment (paragraph splits/merges remain coarse).
- Robust handling of paragraphs that already contain tracked revisions (`<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`) in minimal mode.
- Tabs/line breaks/fields/hyperlinks/content controls as first-class diff tokens (we can add later).

## UX Rules (Must-Haves)

These rules define “good legal diffs” for this feature.

### R1 — Word-level hunks (not character-level)

Compute diffs at token granularity (words/whitespace/punctuation). Do not emit per-character hunks except as a fallback for extremely small tokens if required.

### R2 — Deletion then insertion ordering

When a replacement occurs, the tracked markup order in the paragraph must be:

1. `<w:del>…</w:del>` (strikeout)
2. `<w:ins>…</w:ins>` (underline)

### R3 — Whitespace-only changes suppressed unless adjacent

Whitespace-only diffs (inserting/deleting/replacing tokens that are *only* whitespace) are **not emitted** unless they are directly adjacent to a non-whitespace change in the same local area.

Implication (chosen behavior): if the *only* difference between two paragraphs is whitespace, minimal mode produces **no changes** for that paragraph, and accepting changes will not yield the modified paragraph’s whitespace exactly. This is intentional for reviewer readability.

### R4 — Punctuation-only changes are allowed standalone

If the only change is punctuation (e.g., `;` → `:`), emit a small tracked change for it. Do not suppress punctuation-only edits.

### R5 — Low fragmentation

If minimal mode would produce “too many” change hunks in a single paragraph (see defaults), fall back to coarse delete+insert for that paragraph.

## Default Tuning Parameters (Ship Defaults)

These are explicit defaults to implement in MVP; they can be surfaced later as knobs.

- `MAX_TRACKED_HUNKS_PER_PARAGRAPH = 8`
  - Count hunks after whitespace-only suppression.
- `TOKENIZER_PATTERN = r"(\\s+|[\\w]+(?:[’'\\-][\\w]+)*|[^\\w\\s])"`
  - Notes:
    - Preserves whitespace tokens.
    - Treats hyphenated and apostrophe words as a single word token (e.g., `non-disclosure`, `party’s`).
    - Keeps punctuation as separate tokens (e.g., `,`, `;`, `(`, `)`).
- `PUNCTUATION_TOKEN = token matches r"[^\\w\\s]"`
- `WHITESPACE_TOKEN = token.isspace()`

## Scope: Where This Changes the Code

### Current behavior hot spots

- `src/python_docx_redline/document.py` `Document.compare_to()`:
  - paragraph-level `SequenceMatcher` alignment.
  - “replace” currently implemented as delete old paragraph(s) + insert new paragraph(s).
- `src/python_docx_redline/document.py` `Document.replace_tracked()`:
  - wraps the entire matched span in `<w:del>` and inserts entire replacement in `<w:ins>`.

### MVP scope

- Add minimal editing support to `compare_to()` for 1:1 paragraph replacements.

### Phase 2 scope (optional)

- Add minimal editing support to `replace_tracked()` for within-span replacements.

## Public API Spec

### 1) `compare_to`

Add a backwards-compatible opt-in flag:

- `Document.compare_to(modified: Document, author: str | None = None, minimal_edits: bool = False) -> int`

Behavior:

- `minimal_edits=False` (default): existing behavior unchanged.
- `minimal_edits=True`:
  - For paragraph-level opcodes:
    - `equal`: no changes
    - `insert`: existing insert paragraph behavior
    - `delete`: existing delete paragraph behavior
    - `replace`:
      - if `(i2 - i1) == 1 and (j2 - j1) == 1` (1:1 paragraph replacement), apply intra-paragraph minimal edits.
      - otherwise fall back to existing coarse behavior for that opcode range.

### 2) `replace_tracked` (Phase 2)

Add opt-in flag:

- `replace_tracked(..., minimal_edits: bool = False)`

When `minimal_edits=True`:

- diff `matched_text` vs `replacement_text` and apply minimal edits inside the matched span, with rules R1–R5.
- If unsafe/too fragmented, fall back to existing delete+insert inside the match span.

## Functional Requirements (Acceptance Criteria)

### A) Minimal edit generation

1. **Small word replacement is small**
   - “net 30 days” → “net 45 days”
   - Must produce a tracked deletion for “30” and a tracked insertion for “45”, not a full sentence/paragraph delete+insert.

2. **Deletion then insertion ordering**
   - In the XML for that paragraph, the `<w:del>` element must appear before the `<w:ins>` element at the same location.

3. **Punctuation-only changes are tracked**
   - “Agreement;” → “Agreement:”
   - Must emit a punctuation-only tracked change (no whole-paragraph replacement).

4. **Whitespace-only changes are suppressed**
   - “Section  2” → “Section 2” (only whitespace change)
   - Must emit no tracked changes in minimal mode for that paragraph.

5. **Whitespace changes adjacent to other edits are preserved**
   - Example: “net 30 days” → “net  45 days” (double space inserted adjacent to number change)
   - The whitespace change must be included as part of the same local change area (or otherwise not silently dropped).

6. **Run fragmentation does not break minimal edits**
   - Original paragraph contains multiple runs with different `w:rPr`.
   - Unchanged runs must remain unchanged; deleted text must preserve original formatting where possible.

7. **Safety fallback**
   - If the paragraph would exceed `MAX_TRACKED_HUNKS_PER_PARAGRAPH`, the system must fall back to coarse delete+insert for that paragraph (and still produce valid OOXML).

### B) Output validity

- Resulting OOXML must remain valid and pass the project’s validation checks.
- No modifications outside tracked changes should be introduced in minimal mode.

## Proposed Technical Design

### 1) Paragraph alignment stays as-is

Continue using existing paragraph-level alignment in `compare_to()` to determine which paragraphs are equal/inserted/deleted/replaced.

Minimal mode changes only how 1:1 `replace` paragraphs are handled.

### 2) Token diff (word/space/punct)

For 1:1 paragraph replacements:

1. Extract `orig_text` and `new_text` as the paragraph’s display text (same text basis used today).
2. Tokenize both with `TOKENIZER_PATTERN`, producing `orig_tokens` and `new_tokens`.
3. Compute `SequenceMatcher(orig_tokens, new_tokens, autojunk=False).get_opcodes()`.

### 3) Legal-style opcode normalization (R2–R4)

Transform opcodes into a list of “edit hunks” that will be applied to OOXML.

Definitions:

- A **hunk** is a contiguous edit action with:
  - a token span in the original (`orig_token_start`, `orig_token_end`)
  - an inserted token string (`insert_text`, possibly empty)
  - a deleted token string (`delete_text`, possibly empty)
  - flags: `is_whitespace_only`, `is_punctuation_only`

Rules:

1. Convert `replace` opcodes into paired `delete_text` + `insert_text` at the same boundary.
2. Classify hunks:
   - `is_whitespace_only`: both `delete_text` and `insert_text` consist only of whitespace tokens.
   - `is_punctuation_only`: both `delete_text` and `insert_text` consist only of punctuation tokens (and optional whitespace).
3. Suppress whitespace-only hunks unless adjacent to a non-whitespace hunk:
   - “Adjacent” means immediately preceding or following in the opcode stream with no intervening non-equal tokens.
4. Keep punctuation-only hunks even when standalone.

### 4) Fragmentation guardrail (R5)

After whitespace suppression:

- If number of remaining hunks `> MAX_TRACKED_HUNKS_PER_PARAGRAPH`, fall back to coarse paragraph replacement for that paragraph.

### 5) Applying hunks to OOXML

Core idea: split runs so that hunk boundaries align to run boundaries, then:

- wrap deleted content in `<w:del>` (preserving original runs where feasible),
- insert `<w:ins>` for inserted text,
- leave equal content untouched.

Implementation approach:

1. Build a linear character map for the paragraph:
   - map from character offsets → `(run, offset)` similar to `TextSearch`, but for the full paragraph.
2. Convert token spans to character spans.
3. Split runs at all character boundaries needed:
   - hunk start/end boundaries
4. For each hunk in left-to-right order:
   - If `delete_text` non-empty:
     - take the run slice corresponding to that char span
     - wrap those runs in `<w:del>` and convert `w:t` → `w:delText`
   - If `insert_text` non-empty:
     - insert `<w:ins>` immediately after the deletion (or at the boundary if no deletion)
     - inserted runs should inherit `w:rPr` from nearest unchanged neighbor:
       - prefer left neighbor; else right; else none
5. Coalesce adjacent runs with identical `w:rPr` that were created by splitting (do not attempt to rewrite unrelated existing runs aggressively).

Note: `TrackedXMLGenerator` currently emits `<w:ins>`/`<w:del>` from strings. Minimal mode will likely need either:

- lxml-based construction utilities that accept a `w:rPr` template to apply to inserted runs, or
- new generator methods that accept a `w:rPr` element to clone into inserted runs.

## Safety & Fallback Rules

Minimal mode must fall back to coarse behavior for a paragraph when any of the following are true:

- The paragraph already contains tracked revisions (`<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`) in its subtree.
- The diff produces `> MAX_TRACKED_HUNKS_PER_PARAGRAPH` hunks (after whitespace suppression).
- The paragraph contains unsupported constructs for MVP tokenization/mapping (fields/hyperlinks/content controls) where a safe edit boundary cannot be computed reliably.

Fallback behavior:

- Use the existing coarse paragraph delete+insert approach for that paragraph.

## Test Plan

Add tests using the repo’s existing “minimal valid OOXML” patterns in `tests/`:

1. `compare_to(minimal_edits=True)` replaces a number inside a sentence with small del+ins.
2. Verify ordering: `<w:del>` precedes `<w:ins>` for a replacement.
3. Punctuation-only change produces a small tracked change.
4. Whitespace-only paragraph difference produces 0 changes for that paragraph.
5. Whitespace adjacent to a word change is preserved (not silently dropped).
6. Fragmented runs: ensure only affected runs are wrapped and formatting is preserved.
7. Guardrail: craft a paragraph with many alternating edits and verify fallback triggers.

## Rollout Plan

1. Ship behind `minimal_edits` flag in `compare_to()` only.
2. Add documentation and examples showing recommended usage for legal redlines.
3. After stabilization, add `minimal_edits` to `replace_tracked()` (Phase 2).

## Developer Task Breakdown (Suggested PR Checklist)

1. Add tokenizer + hunk normalization logic (token diff → hunks) with unit tests.
2. Add paragraph linearization + run splitting utilities.
3. Add “apply hunks to paragraph” OOXML implementation with unit tests.
4. Integrate into `Document.compare_to(..., minimal_edits=True)` for 1:1 replacements.
5. Add fallback detection for unsafe paragraphs and fragmentation threshold.
6. (Optional) Phase 2: extend to `replace_tracked(minimal_edits=True)`.
