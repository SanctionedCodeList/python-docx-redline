/**
 * DocTree YAML Serializer
 *
 * Serializes AccessibilityTree to YAML format with three verbosity levels.
 * Based on docs/DOCTREE_SPEC.md section 4 (YAML Output Format).
 *
 * Verbosity Levels:
 * - minimal: Structure overview only (~2000 tokens) - refs and brief content
 * - standard: Full content with styles (default) - complete text and metadata
 * - full: Run-level detail with all formatting information
 */

import type {
  AccessibilityTree,
  AccessibilityNode,
  VerbosityLevel,
  TrackedChange,
  Comment,
  NodeStyle,
  ImageInfo,
  LinkInfo,
  BookmarkInfo,
  FootnoteInfo,
} from './types';
import { SemanticRole } from './types';

// =============================================================================
// State Formatting Helpers
// =============================================================================

/**
 * Build state annotations in [bracket] format.
 *
 * States represent visual/semantic properties like [bold], [italic], [header], [level=2]
 */
function buildStates(node: AccessibilityNode): string[] {
  const states: string[] = [];

  // Add ref as first state
  states.push(`ref=${node.ref}`);

  // Heading level
  if (node.level !== undefined) {
    states.push(`level=${node.level}`);
  }

  // Table dimensions
  if (node.dimensions) {
    states.push(`rows=${node.dimensions.rows}`);
    states.push(`cols=${node.dimensions.cols}`);
  }

  // Header row
  if (node.isHeader) {
    states.push('header');
  }

  // Tracked changes
  if (node.hasChanges) {
    states.push('has-changes');
  }

  // Comments
  if (node.hasComments) {
    states.push('has-comments');
  }

  // Footnotes
  if (node.footnoteRefs?.length) {
    states.push(`fn=${node.footnoteRefs.join(',')}`);
  }

  // Formatting states (for full verbosity)
  if (node.style?.formatting) {
    const fmt = node.style.formatting;
    if (fmt.bold) states.push('bold');
    if (fmt.italic) states.push('italic');
    if (fmt.underline) states.push('underline');
    if (fmt.strikethrough) states.push('strikethrough');
  }

  return states;
}

/**
 * Format states as [state1] [state2] string.
 */
function formatStates(states: string[]): string {
  if (states.length === 0) return '';
  return states.map((s) => `[${s}]`).join(' ');
}

// =============================================================================
// Role Name Formatting
// =============================================================================

/**
 * Get short role name for minimal output.
 */
function getShortRoleName(role: SemanticRole, level?: number): string {
  switch (role) {
    case SemanticRole.Heading:
      return `h${level ?? 1}`;
    case SemanticRole.Paragraph:
      return 'p';
    case SemanticRole.Table:
      return 'table';
    case SemanticRole.Row:
      return 'row';
    case SemanticRole.Cell:
      return 'cell';
    case SemanticRole.ListItem:
      return 'li';
    case SemanticRole.List:
      return 'list';
    case SemanticRole.Blockquote:
      return 'quote';
    case SemanticRole.Image:
      return 'img';
    case SemanticRole.Link:
      return 'link';
    case SemanticRole.Insertion:
      return 'ins';
    case SemanticRole.Deletion:
      return 'del';
    case SemanticRole.Comment:
      return 'comment';
    default:
      return role;
  }
}

/**
 * Get full role name for standard/full output.
 */
function getFullRoleName(role: SemanticRole): string {
  return role; // SemanticRole enum values are already lowercase strings
}

// =============================================================================
// Text Truncation
// =============================================================================

/**
 * Truncate text for minimal output.
 */
function truncateText(text: string, maxLength: number = 50): string {
  if (!text) return '';
  const cleaned = text.replace(/\s+/g, ' ').trim();
  if (cleaned.length <= maxLength) return cleaned;
  return cleaned.substring(0, maxLength - 3) + '...';
}

/**
 * Escape special YAML characters in text.
 */
function escapeYamlText(text: string): string {
  if (!text) return '""';

  // Check if we need quotes
  const needsQuotes =
    text.includes(':') ||
    text.includes('#') ||
    text.includes('\n') ||
    text.includes('"') ||
    text.includes("'") ||
    text.startsWith(' ') ||
    text.endsWith(' ') ||
    text.startsWith('[') ||
    text.startsWith('{') ||
    text === 'true' ||
    text === 'false' ||
    text === 'null' ||
    text === '';

  if (!needsQuotes) return text;

  // Use double quotes and escape internal quotes and newlines
  const escaped = text
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t');

  return `"${escaped}"`;
}

// =============================================================================
// Minimal Verbosity Serialization
// =============================================================================

/**
 * Serialize a node in minimal format.
 * Example: - h1 "SERVICES AGREEMENT" [ref=p:0]
 */
function serializeNodeMinimal(node: AccessibilityNode, indent: string): string {
  const lines: string[] = [];
  const roleName = getShortRoleName(node.role, node.level);

  // Build the main line
  let line = `${indent}- ${roleName}`;

  // Add truncated text for leaf nodes
  if (node.text && !node.children?.length) {
    const truncated = truncateText(node.text, 40);
    line += ` ${escapeYamlText(truncated)}`;
  }

  // Add table summary
  if (node.role === SemanticRole.Table && node.dimensions) {
    line += ` [${node.dimensions.rows}x${node.dimensions.cols}]`;
  }

  // Add ref
  line += ` [ref=${node.ref}]`;

  // Add key states
  if (node.isHeader) line += ' [header]';
  if (node.hasChanges) line += ' [has-changes]';

  lines.push(line);

  // Process children (but limit depth for tables)
  if (node.children && node.role !== SemanticRole.Table) {
    for (const child of node.children) {
      lines.push(serializeNodeMinimal(child, indent + '  '));
    }
  }

  return lines.join('\n');
}

// =============================================================================
// Standard Verbosity Serialization
// =============================================================================

/**
 * Serialize style information.
 */
function serializeStyle(style: NodeStyle, indent: string): string {
  const lines: string[] = [];

  if (style.name) {
    lines.push(`${indent}style: ${style.name}`);
  }

  return lines.join('\n');
}

/**
 * Serialize a node in standard format.
 */
function serializeNodeStandard(node: AccessibilityNode, indent: string): string {
  const lines: string[] = [];
  const roleName = getFullRoleName(node.role);
  const states = buildStates(node);
  const stateStr = formatStates(states);

  // For simple leaf nodes (paragraphs with just text), use inline format
  if (
    node.text &&
    !node.children?.length &&
    !node.style?.name &&
    !node.hasChanges &&
    !node.hasComments
  ) {
    lines.push(`${indent}- ${roleName} ${stateStr}: ${escapeYamlText(node.text)}`);
    return lines.join('\n');
  }

  // For complex nodes, use block format
  lines.push(`${indent}- ${roleName} ${stateStr}:`);

  // Add text content
  if (node.text) {
    lines.push(`${indent}    text: ${escapeYamlText(node.text)}`);
  }

  // Add style
  if (node.style?.name) {
    lines.push(`${indent}    style: ${node.style.name}`);
  }

  // Add tracked change refs
  if (node.changeRefs?.length) {
    lines.push(`${indent}    change_refs: [${node.changeRefs.join(', ')}]`);
  }

  // Add comment refs
  if (node.commentRefs?.length) {
    lines.push(`${indent}    comment_refs: [${node.commentRefs.join(', ')}]`);
  }

  // Process children
  if (node.children?.length) {
    for (const child of node.children) {
      // For table cells, use simplified format
      if (node.role === SemanticRole.Row && child.role === SemanticRole.Cell) {
        if (child.text) {
          lines.push(`${indent}    - cell: ${escapeYamlText(child.text)}`);
        } else if (child.children?.length) {
          lines.push(`${indent}    - cell [ref=${child.ref}]:`);
          for (const cellChild of child.children) {
            lines.push(serializeNodeStandard(cellChild, indent + '        '));
          }
        } else {
          lines.push(`${indent}    - cell: ""`);
        }
      } else {
        lines.push(serializeNodeStandard(child, indent + '    '));
      }
    }
  }

  return lines.join('\n');
}

// =============================================================================
// Full Verbosity Serialization
// =============================================================================

/**
 * Serialize formatting details.
 */
function serializeFormatting(
  formatting: NodeStyle['formatting'],
  indent: string
): string {
  if (!formatting) return '';

  const lines: string[] = [];
  lines.push(`${indent}formatting:`);

  if (formatting.bold) lines.push(`${indent}  bold: true`);
  if (formatting.italic) lines.push(`${indent}  italic: true`);
  if (formatting.underline) lines.push(`${indent}  underline: true`);
  if (formatting.strikethrough) lines.push(`${indent}  strikethrough: true`);
  if (formatting.font) lines.push(`${indent}  font: ${escapeYamlText(formatting.font)}`);
  if (formatting.size) lines.push(`${indent}  size: ${formatting.size}`);
  if (formatting.color) lines.push(`${indent}  color: ${formatting.color}`);
  if (formatting.highlight) lines.push(`${indent}  highlight: ${formatting.highlight}`);

  return lines.length > 1 ? lines.join('\n') : '';
}

/**
 * Serialize run-level detail for full verbosity.
 */
function serializeRuns(runs: AccessibilityNode[], indent: string): string {
  const lines: string[] = [];
  lines.push(`${indent}runs:`);

  for (const run of runs) {
    const states: string[] = [`ref=${run.ref}`];

    // Add formatting states
    if (run.style?.formatting) {
      const fmt = run.style.formatting;
      if (fmt.bold) states.push('bold');
      if (fmt.italic) states.push('italic');
      if (fmt.underline) states.push('underline');
    }

    // Add tracked change state
    if (run.role === SemanticRole.Insertion) states.push('tracked-insert');
    if (run.role === SemanticRole.Deletion) states.push('tracked-delete');

    const stateStr = formatStates(states);
    lines.push(`${indent}  - text ${escapeYamlText(run.text ?? '')} ${stateStr}`);
  }

  return lines.join('\n');
}

/**
 * Serialize a node in full format with run-level detail.
 */
function serializeNodeFull(node: AccessibilityNode, indent: string): string {
  const lines: string[] = [];
  const roleName = getFullRoleName(node.role);
  const states = buildStates(node);
  const stateStr = formatStates(states);

  lines.push(`${indent}- ${roleName} ${stateStr}:`);

  // Add style name
  if (node.style?.name) {
    lines.push(`${indent}    style: ${node.style.name}`);
  }

  // Add formatting details
  if (node.style?.formatting) {
    const fmtLines = serializeFormatting(node.style.formatting, indent + '    ');
    if (fmtLines) lines.push(fmtLines);
  }

  // Add runs if available (full verbosity)
  if (node.runs?.length) {
    lines.push(serializeRuns(node.runs, indent + '    '));
  } else if (node.text) {
    // Fall back to text if no runs
    lines.push(`${indent}    text: ${escapeYamlText(node.text)}`);
  }

  // Add tracked change refs
  if (node.changeRefs?.length) {
    lines.push(`${indent}    change_refs: [${node.changeRefs.join(', ')}]`);
  }

  // Add comment refs
  if (node.commentRefs?.length) {
    lines.push(`${indent}    comment_refs: [${node.commentRefs.join(', ')}]`);
  }

  // Add images
  if (node.images?.length) {
    lines.push(`${indent}    images:`);
    for (const img of node.images) {
      lines.push(serializeImage(img, indent + '      ', 'full'));
    }
  }

  // Process children
  if (node.children?.length) {
    for (const child of node.children) {
      lines.push(serializeNodeFull(child, indent + '    '));
    }
  }

  return lines.join('\n');
}

// =============================================================================
// Auxiliary Content Serialization
// =============================================================================

/**
 * Serialize image information.
 */
function serializeImage(
  img: ImageInfo,
  indent: string,
  verbosity: VerbosityLevel
): string {
  const lines: string[] = [];
  const posType = img.positionType === 'floating' ? '[floating]' : '[inline]';

  lines.push(`${indent}- image [ref=${img.ref}] ${posType}:`);
  if (img.name) lines.push(`${indent}    name: ${escapeYamlText(img.name)}`);
  if (img.altText) lines.push(`${indent}    alt_text: ${escapeYamlText(img.altText)}`);

  if (verbosity === 'full') {
    if (img.sizeEmu) {
      lines.push(`${indent}    size:`);
      lines.push(`${indent}      width_emu: ${img.sizeEmu.widthEmu}`);
      lines.push(`${indent}      height_emu: ${img.sizeEmu.heightEmu}`);
    }
    if (img.format) lines.push(`${indent}    format: ${img.format}`);
    if (img.relationshipId) lines.push(`${indent}    relationship_id: ${img.relationshipId}`);
  } else if (img.size) {
    lines.push(`${indent}    size: ${escapeYamlText(img.size)}`);
  }

  return lines.join('\n');
}

/**
 * Serialize tracked changes summary.
 */
function serializeTrackedChanges(changes: TrackedChange[], indent: string): string {
  const lines: string[] = [];
  lines.push(`${indent}tracked_changes:`);

  for (const change of changes) {
    lines.push(`${indent}  - ref: ${change.ref}`);
    lines.push(`${indent}    type: ${change.type === 'ins' ? 'insertion' : 'deletion'}`);
    if (change.id) lines.push(`${indent}    id: ${escapeYamlText(change.id)}`);
    lines.push(`${indent}    author: ${escapeYamlText(change.author)}`);
    lines.push(`${indent}    date: ${escapeYamlText(change.date)}`);
    lines.push(`${indent}    text: ${escapeYamlText(change.text)}`);
    if (change.location?.paragraphRef) {
      lines.push(`${indent}    location: ${change.location.paragraphRef}`);
    }
  }

  return lines.join('\n');
}

/**
 * Serialize comments summary.
 */
function serializeComments(comments: Comment[], indent: string): string {
  const lines: string[] = [];
  lines.push(`${indent}comments:`);

  for (const comment of comments) {
    lines.push(`${indent}  - ref: ${comment.ref}`);
    lines.push(`${indent}    author: ${escapeYamlText(comment.author)}`);
    lines.push(`${indent}    text: ${escapeYamlText(comment.text)}`);
    if (comment.onText) lines.push(`${indent}    on_text: ${escapeYamlText(comment.onText)}`);
    lines.push(`${indent}    resolved: ${comment.resolved}`);

    if (comment.replies?.length) {
      lines.push(`${indent}    replies:`);
      for (const reply of comment.replies) {
        lines.push(`${indent}      - ref: ${reply.ref}`);
        lines.push(`${indent}        author: ${escapeYamlText(reply.author)}`);
        lines.push(`${indent}        text: ${escapeYamlText(reply.text)}`);
      }
    }
  }

  return lines.join('\n');
}

/**
 * Serialize bookmarks summary.
 */
function serializeBookmarks(bookmarks: BookmarkInfo[], indent: string): string {
  const lines: string[] = [];
  lines.push(`${indent}bookmarks:`);

  for (const bk of bookmarks) {
    lines.push(`${indent}  - ref: ${bk.ref}`);
    lines.push(`${indent}    name: ${escapeYamlText(bk.name)}`);
    lines.push(`${indent}    location: ${bk.location}`);
    if (bk.textPreview) {
      lines.push(`${indent}    text_preview: ${escapeYamlText(bk.textPreview)}`);
    }
    if (bk.referencedBy?.length) {
      lines.push(`${indent}    referenced_by: [${bk.referencedBy.join(', ')}]`);
    }
  }

  return lines.join('\n');
}

/**
 * Serialize footnotes summary.
 */
function serializeFootnotes(footnotes: FootnoteInfo[], indent: string): string {
  const lines: string[] = [];
  lines.push(`${indent}footnotes:`);

  for (const fn of footnotes) {
    lines.push(`${indent}  - ref: ${fn.ref}`);
    lines.push(`${indent}    id: ${fn.id}`);
    lines.push(`${indent}    text: ${escapeYamlText(fn.text)}`);
    if (fn.referencedFrom) {
      lines.push(`${indent}    referenced_from: ${fn.referencedFrom}`);
    }
  }

  return lines.join('\n');
}

// =============================================================================
// Main Serialization Function
// =============================================================================

/**
 * Serialize an AccessibilityTree to YAML format.
 *
 * @param tree - The accessibility tree to serialize
 * @param verbosity - Output verbosity level (minimal, standard, full)
 * @returns YAML string representation
 *
 * @example
 * ```typescript
 * // Minimal output (~2000 tokens for navigation)
 * const minimal = treeToYaml(tree, 'minimal');
 *
 * // Standard output (default, full content)
 * const standard = treeToYaml(tree);
 *
 * // Full output (run-level detail)
 * const full = treeToYaml(tree, 'full');
 * ```
 */
export function treeToYaml(
  tree: AccessibilityTree,
  verbosity: VerbosityLevel = 'standard'
): string {
  const lines: string[] = [];

  // Document metadata header
  lines.push('document:');
  if (tree.document.path) {
    lines.push(`  path: ${escapeYamlText(tree.document.path)}`);
  }
  lines.push(`  verbosity: ${verbosity}`);

  // Stats
  lines.push('  stats:');
  lines.push(`    paragraphs: ${tree.document.stats.paragraphs}`);
  lines.push(`    tables: ${tree.document.stats.tables}`);
  lines.push(`    tracked_changes: ${tree.document.stats.trackedChanges}`);
  lines.push(`    comments: ${tree.document.stats.comments}`);
  if (tree.document.stats.footnotes !== undefined) {
    lines.push(`    footnotes: ${tree.document.stats.footnotes}`);
  }
  if (tree.document.stats.sections !== undefined) {
    lines.push(`    sections: ${tree.document.stats.sections}`);
  }

  // Section detection info (for outline mode)
  if (tree.document.sectionDetection) {
    lines.push('  section_detection:');
    lines.push(`    method: ${tree.document.sectionDetection.method}`);
    lines.push(`    confidence: ${tree.document.sectionDetection.confidence}`);
  }

  lines.push('');

  // Content based on verbosity
  const nodes = tree.content ?? tree.outline ?? [];

  if (verbosity === 'minimal') {
    // Minimal: outline format
    if (nodes.length > 0) {
      lines.push('outline:');
      for (const node of nodes) {
        lines.push(serializeNodeMinimal(node, '  '));
      }
    }
  } else if (verbosity === 'standard') {
    // Standard: content format
    if (nodes.length > 0) {
      lines.push('content:');
      for (const node of nodes) {
        lines.push(serializeNodeStandard(node, '  '));
      }
    }
  } else {
    // Full: detailed format
    if (nodes.length > 0) {
      lines.push('content:');
      for (const node of nodes) {
        lines.push(serializeNodeFull(node, '  '));
      }
    }
  }

  // Tracked changes summary (standard and full)
  if (verbosity !== 'minimal' && tree.trackedChanges?.length) {
    lines.push('');
    lines.push(serializeTrackedChanges(tree.trackedChanges, ''));
  }

  // Comments summary (standard and full)
  if (verbosity !== 'minimal' && tree.comments?.length) {
    lines.push('');
    lines.push(serializeComments(tree.comments, ''));
  }

  // Footnotes (all verbosity levels - important for legal docs)
  if (tree.footnotes?.length) {
    lines.push('');
    lines.push(serializeFootnotes(tree.footnotes, ''));
  }

  // Bookmarks (full only)
  if (verbosity === 'full' && tree.bookmarks?.length) {
    lines.push('');
    lines.push(serializeBookmarks(tree.bookmarks, ''));
  }

  // Navigation hints (minimal only)
  if (verbosity === 'minimal' && tree.navigation) {
    lines.push('');
    lines.push('navigation:');
    lines.push(`  expand_section: ${escapeYamlText(tree.navigation.expandSection)}`);
    lines.push(`  search: ${escapeYamlText(tree.navigation.search)}`);
  }

  return lines.join('\n');
}

// =============================================================================
// Convenience Functions
// =============================================================================

/**
 * Serialize tree with minimal verbosity (~2000 tokens target).
 */
export function toMinimalYaml(tree: AccessibilityTree): string {
  return treeToYaml(tree, 'minimal');
}

/**
 * Serialize tree with standard verbosity (default).
 */
export function toStandardYaml(tree: AccessibilityTree): string {
  return treeToYaml(tree, 'standard');
}

/**
 * Serialize tree with full verbosity (run-level detail).
 */
export function toFullYaml(tree: AccessibilityTree): string {
  return treeToYaml(tree, 'full');
}
