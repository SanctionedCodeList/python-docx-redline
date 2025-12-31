/**
 * DocTree Tracked Changes and Comments Support
 *
 * Extends the accessibility tree to detect and include:
 * - Tracked insertions (w:ins in OOXML)
 * - Tracked deletions (w:del in OOXML)
 * - Comments with text and author
 * - Comment threading (replies)
 *
 * Based on docs/DOCTREE_SPEC.md section 5 (Track Changes & Comments).
 *
 * Note: Office.js has limited tracked changes API. We use:
 * - body.getReviewedText() for different change views
 * - body.getComments() for comments (if available)
 * - Paragraph text comparison to detect changes
 */

import type {
  TrackedChange,
  TrackedChangeType,
  Comment,
  CommentReply,
  AccessibilityNode,
  ChangeViewMode,
  Ref,
} from './types';

// =============================================================================
// Office.js Type Declarations for Tracked Changes
// =============================================================================

/**
 * Minimal Office.js Word.RequestContext interface for tracked changes.
 * Must match the interface in builder.ts for compatibility.
 */
interface WordRequestContext {
  document: {
    body: {
      paragraphs: WordParagraphCollection;
      tables?: unknown;
      getReviewedText(changeTrackingVersion: string): WordStringResult;
      getComments(): WordCommentCollection;
    };
  };
  sync(): Promise<void>;
}

interface WordParagraphCollection {
  load(properties: string): WordParagraphCollection;
  items: WordParagraph[];
}

interface WordParagraph {
  text: string;
  style: string;
  getReviewedText(changeTrackingVersion: string): WordStringResult;
}

interface WordRange {
  text: string;
  load(properties: string): WordRange;
}

interface WordStringResult {
  value: string;
}

interface WordCommentCollection {
  load(properties: string): WordCommentCollection;
  items: WordComment[];
}

interface WordComment {
  id: string;
  authorName: string;
  authorEmail: string;
  content: string;
  createdDate: Date;
  resolved: boolean;
  replies: WordCommentReplyCollection;
  getRange(): WordRange;
  load(properties: string): WordComment;
}

interface WordCommentReplyCollection {
  load(properties: string): WordCommentReplyCollection;
  items: WordCommentReply[];
}

interface WordCommentReply {
  id: string;
  authorName: string;
  authorEmail: string;
  content: string;
  createdDate: Date;
}

// =============================================================================
// Change View Mode Mapping
// =============================================================================

/**
 * Maps our ChangeViewMode to Office.js ChangeTrackingVersion values.
 *
 * Office.js uses:
 * - 'Current' - shows the document with all changes applied (final)
 * - 'Original' - shows the document before changes (original)
 *
 * Note: There's no direct 'markup' mode in getReviewedText - we detect
 * changes by comparing Original and Current versions.
 */
type OfficeChangeTrackingVersion = 'Current' | 'Original';

function getOfficeVersion(mode: ChangeViewMode): OfficeChangeTrackingVersion {
  switch (mode) {
    case 'final':
      return 'Current';
    case 'original':
      return 'Original';
    case 'markup':
      // For markup mode, we'll compare both versions
      return 'Current';
  }
}

// =============================================================================
// Tracked Change Detection
// =============================================================================

/**
 * Result of detecting tracked changes in a paragraph.
 */
export interface ParagraphChanges {
  /** Original text (before changes) */
  originalText: string;
  /** Current text (after changes) */
  currentText: string;
  /** Whether this paragraph has tracked changes */
  hasChanges: boolean;
  /** Detected tracked changes */
  changes: TrackedChange[];
  /** Text with inline change markers for markup mode */
  markedUpText: string;
}

/**
 * Generate a tracked change ref.
 */
function makeChangeRef(index: number): Ref {
  return `change:${index}`;
}

/**
 * Generate a comment ref.
 */
function makeCommentRef(index: number): Ref {
  return `cmt:${index}`;
}

/**
 * Detect changes by comparing original and current text.
 *
 * This uses a simple diff algorithm to find insertions and deletions.
 * For more precise detection, we'd need access to the raw OOXML.
 */
export function detectChanges(
  originalText: string,
  currentText: string,
  paragraphRef: Ref,
  changeIndexStart: number,
  author: string = 'Unknown',
  date: string = new Date().toISOString()
): { changes: TrackedChange[]; markedUpText: string } {
  const changes: TrackedChange[] = [];

  // If texts are identical, no changes
  if (originalText === currentText) {
    return { changes: [], markedUpText: currentText };
  }

  // Simple word-level diff
  const originalWords = originalText.split(/\s+/).filter((w) => w.length > 0);
  const currentWords = currentText.split(/\s+/).filter((w) => w.length > 0);

  let changeIndex = changeIndexStart;
  const markedUpParts: string[] = [];

  // Use longest common subsequence (LCS) approach for better diff
  const lcs = findLCS(originalWords, currentWords);

  let origIdx = 0;
  let currIdx = 0;
  let lcsIdx = 0;

  while (origIdx < originalWords.length || currIdx < currentWords.length) {
    const lcsWord = lcsIdx < lcs.length ? lcs[lcsIdx] : null;

    // Check for deletions (in original but not in current at this position)
    if (
      origIdx < originalWords.length &&
      originalWords[origIdx] !== lcsWord
    ) {
      const deletedWord = originalWords[origIdx];
      if (deletedWord) {
        changes.push({
          ref: makeChangeRef(changeIndex++),
          type: 'del',
          author,
          date,
          text: deletedWord,
          location: { paragraphRef },
        });
        markedUpParts.push(`{--${deletedWord}--}`);
      }
      origIdx++;
      continue;
    }

    // Check for insertions (in current but not in original at this position)
    if (
      currIdx < currentWords.length &&
      currentWords[currIdx] !== lcsWord
    ) {
      const insertedWord = currentWords[currIdx];
      if (insertedWord) {
        changes.push({
          ref: makeChangeRef(changeIndex++),
          type: 'ins',
          author,
          date,
          text: insertedWord,
          location: { paragraphRef },
        });
        markedUpParts.push(`{++${insertedWord}++}`);
      }
      currIdx++;
      continue;
    }

    // Common word - advance all pointers
    if (lcsWord && currentWords[currIdx] === lcsWord) {
      markedUpParts.push(lcsWord);
      origIdx++;
      currIdx++;
      lcsIdx++;
    } else {
      // Safety: advance to avoid infinite loop
      if (currIdx < currentWords.length) {
        const word = currentWords[currIdx];
        if (word) markedUpParts.push(word);
        currIdx++;
      }
      if (origIdx < originalWords.length) origIdx++;
    }
  }

  return {
    changes,
    markedUpText: markedUpParts.join(' '),
  };
}

/**
 * Find longest common subsequence of two word arrays.
 */
function findLCS(a: string[], b: string[]): string[] {
  const m = a.length;
  const n = b.length;

  // Create LCS table
  const dp: number[][] = Array(m + 1)
    .fill(null)
    .map(() => Array(n + 1).fill(0) as number[]);

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (a[i - 1] === b[j - 1]) {
        dp[i]![j] = (dp[i - 1]?.[j - 1] ?? 0) + 1;
      } else {
        dp[i]![j] = Math.max(dp[i - 1]?.[j] ?? 0, dp[i]?.[j - 1] ?? 0);
      }
    }
  }

  // Backtrack to find LCS
  const lcs: string[] = [];
  let i = m;
  let j = n;
  while (i > 0 && j > 0) {
    if (a[i - 1] === b[j - 1]) {
      const word = a[i - 1];
      if (word) lcs.unshift(word);
      i--;
      j--;
    } else if ((dp[i - 1]?.[j] ?? 0) > (dp[i]?.[j - 1] ?? 0)) {
      i--;
    } else {
      j--;
    }
  }

  return lcs;
}

/**
 * Collect tracked changes for all paragraphs.
 *
 * This function:
 * 1. Gets original and current text for each paragraph
 * 2. Compares to detect insertions/deletions
 * 3. Returns change info for each paragraph and overall change list
 *
 * PERFORMANCE: Uses batched context.sync() calls to avoid O(n) round-trips.
 * Previously called context.sync() per paragraph, causing 3+ minute delays
 * for large documents.
 */
export async function collectTrackedChanges(
  context: WordRequestContext,
  paragraphRefs: Ref[],
  mode: ChangeViewMode = 'markup'
): Promise<{
  paragraphChanges: Map<Ref, ParagraphChanges>;
  allChanges: TrackedChange[];
}> {
  const paragraphChanges = new Map<Ref, ParagraphChanges>();
  const allChanges: TrackedChange[] = [];
  let changeIndex = 0;

  // Load paragraphs
  const paragraphs = context.document.body.paragraphs.load('items/text,items/style');
  await context.sync();

  // BATCHED APPROACH: Queue all getReviewedText calls, then sync once
  // This reduces O(n) sync calls to O(1)
  const reviewedTextResults: Array<{
    ref: Ref;
    paraText: string;
    originalResult: WordStringResult;
    currentResult: WordStringResult;
  }> = [];

  // Phase 1: Queue all getReviewedText calls (no sync yet)
  for (let i = 0; i < paragraphs.items.length && i < paragraphRefs.length; i++) {
    const para = paragraphs.items[i];
    const ref = paragraphRefs[i];

    if (!para || !ref) continue;

    try {
      // Queue both version requests (Office.js batches these)
      const originalResult = para.getReviewedText('Original');
      const currentResult = para.getReviewedText('Current');

      reviewedTextResults.push({
        ref,
        paraText: para.text,
        originalResult,
        currentResult,
      });
    } catch {
      // getReviewedText may not be available - use plain text
      paragraphChanges.set(ref, {
        originalText: para.text,
        currentText: para.text,
        hasChanges: false,
        changes: [],
        markedUpText: para.text,
      });
    }
  }

  // Phase 2: Single sync for all batched operations
  if (reviewedTextResults.length > 0) {
    try {
      await context.sync();
    } catch {
      // If batched sync fails, fall back to plain text for all
      for (const item of reviewedTextResults) {
        paragraphChanges.set(item.ref, {
          originalText: item.paraText,
          currentText: item.paraText,
          hasChanges: false,
          changes: [],
          markedUpText: item.paraText,
        });
      }
      return { paragraphChanges, allChanges };
    }
  }

  // Phase 3: Process all results (synchronous - no more API calls)
  for (const item of reviewedTextResults) {
    try {
      const originalText = item.originalResult.value;
      const currentText = item.currentResult.value;
      const hasChanges = originalText !== currentText;

      let changes: TrackedChange[] = [];
      let markedUpText = currentText;

      if (hasChanges && mode === 'markup') {
        const detected = detectChanges(
          originalText,
          currentText,
          item.ref,
          changeIndex
        );
        changes = detected.changes;
        markedUpText = detected.markedUpText;
        changeIndex += changes.length;
        allChanges.push(...changes);
      }

      paragraphChanges.set(item.ref, {
        originalText,
        currentText,
        hasChanges,
        changes,
        markedUpText,
      });
    } catch {
      // Individual result access failed - use plain text
      paragraphChanges.set(item.ref, {
        originalText: item.paraText,
        currentText: item.paraText,
        hasChanges: false,
        changes: [],
        markedUpText: item.paraText,
      });
    }
  }

  return { paragraphChanges, allChanges };
}

// =============================================================================
// Comment Collection
// =============================================================================

/**
 * Collect all comments from the document.
 *
 * PERFORMANCE: Uses batched context.sync() calls to minimize round-trips.
 * Previously called context.sync() per comment for replies and ranges.
 */
export async function collectComments(
  context: WordRequestContext
): Promise<Comment[]> {
  const comments: Comment[] = [];

  try {
    // Phase 1: Load comments
    const wordComments = context.document.body.getComments();
    wordComments.load('items/id,items/authorName,items/content,items/createdDate,items/resolved');
    await context.sync();

    if (wordComments.items.length === 0) {
      return comments;
    }

    // Phase 2: Queue loading all replies (batched)
    for (const wc of wordComments.items) {
      try {
        wc.replies.load('items/id,items/authorName,items/content,items/createdDate');
      } catch {
        // Replies may not be supported for this comment
      }
    }

    // Phase 3: Queue loading all ranges (batched)
    const ranges: WordRange[] = [];
    for (const wc of wordComments.items) {
      try {
        const range = wc.getRange();
        range.load('text');
        ranges.push(range);
      } catch {
        // Range may not be available - push null placeholder
        ranges.push(null as unknown as WordRange);
      }
    }

    // Single sync for all replies and ranges
    await context.sync();

    // Phase 4: Process all results (synchronous)
    let commentIndex = 0;
    for (let i = 0; i < wordComments.items.length; i++) {
      const wc = wordComments.items[i];
      if (!wc) continue;

      // Process replies
      let replies: CommentReply[] = [];
      try {
        let replyIndex = 0;
        for (const wr of wc.replies.items) {
          replies.push({
            ref: `${makeCommentRef(commentIndex)}/reply:${replyIndex}`,
            author: wr.authorName || 'Unknown',
            date: wr.createdDate?.toISOString() || new Date().toISOString(),
            text: wr.content || '',
          });
          replyIndex++;
        }
      } catch {
        // Replies access failed
      }

      // Get range text
      let onText: string | undefined;
      try {
        const range = ranges[i];
        if (range) {
          onText = range.text;
        }
      } catch {
        // Range access failed
      }

      comments.push({
        ref: makeCommentRef(commentIndex),
        id: wc.id,
        author: wc.authorName || 'Unknown',
        date: wc.createdDate?.toISOString() || new Date().toISOString(),
        text: wc.content || '',
        onText,
        resolved: wc.resolved || false,
        replies: replies.length > 0 ? replies : undefined,
      });

      commentIndex++;
    }
  } catch {
    // Comments API may not be available
    // Return empty array - this is expected in some environments
  }

  return comments;
}

// =============================================================================
// Integration with AccessibilityNode
// =============================================================================

/**
 * Apply tracked change information to an AccessibilityNode.
 */
export function applyChangesToNode(
  node: AccessibilityNode,
  paragraphChanges: Map<Ref, ParagraphChanges>,
  mode: ChangeViewMode
): AccessibilityNode {
  const changes = paragraphChanges.get(node.ref);

  if (!changes) {
    return node;
  }

  // Apply text based on view mode
  let text: string;
  switch (mode) {
    case 'final':
      text = changes.currentText;
      break;
    case 'original':
      text = changes.originalText;
      break;
    case 'markup':
      text = changes.markedUpText;
      break;
  }

  return {
    ...node,
    text,
    hasChanges: changes.hasChanges,
    changes: changes.changes.length > 0 ? changes.changes : undefined,
    changeRefs: changes.changes.length > 0
      ? changes.changes.map((c) => c.ref)
      : undefined,
  };
}

/**
 * Apply comments to an AccessibilityNode based on text matching.
 */
export function applyCommentsToNode(
  node: AccessibilityNode,
  comments: Comment[]
): AccessibilityNode {
  // Find comments that reference text in this node
  const nodeComments = comments.filter((c) => {
    if (!c.onText || !node.text) return false;
    return node.text.includes(c.onText);
  });

  if (nodeComments.length === 0) {
    return node;
  }

  return {
    ...node,
    hasComments: true,
    comments: nodeComments,
    commentRefs: nodeComments.map((c) => c.ref),
  };
}

// =============================================================================
// Export Helpers
// =============================================================================

/**
 * Options for change and comment processing.
 */
export interface ChangeProcessingOptions {
  /** Change view mode */
  changeViewMode: ChangeViewMode;
  /** Include tracked changes */
  includeTrackedChanges: boolean;
  /** Include comments */
  includeComments: boolean;
}

/**
 * Result of change and comment processing.
 */
export interface ChangeProcessingResult {
  /** All tracked changes */
  trackedChanges: TrackedChange[];
  /** All comments */
  comments: Comment[];
  /** Paragraph changes map */
  paragraphChanges: Map<Ref, ParagraphChanges>;
  /** Stats update */
  stats: {
    trackedChanges: number;
    comments: number;
  };
}

/**
 * Process tracked changes and comments for a document.
 */
export async function processChangesAndComments(
  context: WordRequestContext,
  paragraphRefs: Ref[],
  options: ChangeProcessingOptions
): Promise<ChangeProcessingResult> {
  let trackedChanges: TrackedChange[] = [];
  let paragraphChanges = new Map<Ref, ParagraphChanges>();
  let comments: Comment[] = [];

  if (options.includeTrackedChanges) {
    const result = await collectTrackedChanges(
      context,
      paragraphRefs,
      options.changeViewMode
    );
    trackedChanges = result.allChanges;
    paragraphChanges = result.paragraphChanges;
  }

  if (options.includeComments) {
    comments = await collectComments(context);
  }

  return {
    trackedChanges,
    comments,
    paragraphChanges,
    stats: {
      trackedChanges: trackedChanges.length,
      comments: comments.length,
    },
  };
}
