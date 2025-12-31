/**
 * Scope System for DocTree Accessibility Layer
 *
 * Provides filtering capabilities for accessibility nodes based on
 * text content, section membership, style, role, and custom predicates.
 *
 * Ported from python-docx-redline scope system.
 */

import type {
  AccessibilityNode,
  AccessibilityTree,
  NoteScope,
  Ref,
  ScopeFilter,
  ScopeOptions,
  ScopePredicate,
  ScopeResult,
  ScopeSpec,
} from './types';
import { SemanticRole } from './types';

// =============================================================================
// String Scope Parsing
// =============================================================================

/**
 * Parse a string scope shortcut into a ScopeFilter.
 *
 * Supported formats:
 * - "keyword"                    → { contains: "keyword" }
 * - "section:Introduction"       → { section: "Introduction" }
 * - "paragraph_containing:text"  → { contains: "text" }
 * - "role:heading"               → { role: SemanticRole.Heading }
 * - "style:Normal"               → { style: "Normal" }
 * - "footnotes"                  → { role: SemanticRole.Footnote }
 * - "footnote:1"                 → { refs: ["fn:1"] }
 * - "endnotes"                   → { role: SemanticRole.Endnote }
 * - "endnote:1"                  → { refs: ["en:1"] }
 *
 * @param scopeString - String scope specification
 * @returns Normalized ScopeFilter
 */
export function parseScope(scopeString: string): ScopeFilter {
  const trimmed = scopeString.trim();

  // Section prefix
  if (trimmed.startsWith('section:')) {
    return { section: trimmed.slice(8) };
  }

  // Paragraph containing prefix (explicit)
  if (trimmed.startsWith('paragraph_containing:')) {
    return { contains: trimmed.slice(21) };
  }

  // Role prefix
  if (trimmed.startsWith('role:')) {
    const roleStr = trimmed.slice(5).toLowerCase();
    const role = parseRole(roleStr);
    if (!role) {
      throw new Error(`Unknown role: ${roleStr}`);
    }
    return { role };
  }

  // Style prefix
  if (trimmed.startsWith('style:')) {
    return { style: trimmed.slice(6) };
  }

  // Footnote/endnote shortcuts
  if (trimmed === 'footnotes') {
    return { role: SemanticRole.Footnote };
  }
  if (trimmed === 'endnotes') {
    return { role: SemanticRole.Endnote };
  }
  if (trimmed === 'notes') {
    return { role: [SemanticRole.Footnote, SemanticRole.Endnote] };
  }
  if (trimmed.startsWith('footnote:')) {
    const id = trimmed.slice(9);
    return { refs: [`fn:${id}` as Ref] };
  }
  if (trimmed.startsWith('endnote:')) {
    const id = trimmed.slice(8);
    return { refs: [`en:${id}` as Ref] };
  }

  // Level prefix (for headings)
  if (trimmed.startsWith('level:')) {
    const level = parseInt(trimmed.slice(6), 10);
    if (isNaN(level)) {
      throw new Error(`Invalid level: ${trimmed.slice(6)}`);
    }
    return { role: SemanticRole.Heading, level };
  }

  // Has changes/comments shortcuts
  if (trimmed === 'changes' || trimmed === 'tracked') {
    return { hasChanges: true };
  }
  if (trimmed === 'comments') {
    return { hasComments: true };
  }

  // Default: text contains
  return { contains: trimmed };
}

/**
 * Parse a role string to SemanticRole enum value.
 */
function parseRole(roleStr: string): SemanticRole | undefined {
  const normalized = roleStr.toLowerCase();
  const roleMap: Record<string, SemanticRole> = {
    document: SemanticRole.Document,
    header: SemanticRole.Header,
    footer: SemanticRole.Footer,
    section: SemanticRole.Section,
    heading: SemanticRole.Heading,
    paragraph: SemanticRole.Paragraph,
    blockquote: SemanticRole.Blockquote,
    list: SemanticRole.List,
    listitem: SemanticRole.ListItem,
    table: SemanticRole.Table,
    row: SemanticRole.Row,
    cell: SemanticRole.Cell,
    text: SemanticRole.Text,
    strong: SemanticRole.Strong,
    emphasis: SemanticRole.Emphasis,
    link: SemanticRole.Link,
    insertion: SemanticRole.Insertion,
    deletion: SemanticRole.Deletion,
    comment: SemanticRole.Comment,
    image: SemanticRole.Image,
    chart: SemanticRole.Chart,
    diagram: SemanticRole.Diagram,
    shape: SemanticRole.Shape,
    footnote: SemanticRole.Footnote,
    endnote: SemanticRole.Endnote,
    bookmark: SemanticRole.Bookmark,
  };
  return roleMap[normalized];
}

// =============================================================================
// Note Scope Parsing
// =============================================================================

/**
 * Parse a scope string to check if it targets footnotes/endnotes.
 *
 * @param scopeString - String scope specification
 * @returns NoteScope if this is a note-related scope, null otherwise
 */
export function parseNoteScope(scopeString: string): NoteScope | null {
  const trimmed = scopeString.trim().toLowerCase();

  if (trimmed === 'footnotes') {
    return { scopeType: 'footnotes' };
  }
  if (trimmed === 'endnotes') {
    return { scopeType: 'endnotes' };
  }
  if (trimmed === 'notes') {
    return { scopeType: 'notes' };
  }
  if (trimmed.startsWith('footnote:')) {
    return { scopeType: 'footnote', noteId: trimmed.slice(9) };
  }
  if (trimmed.startsWith('endnote:')) {
    return { scopeType: 'endnote', noteId: trimmed.slice(8) };
  }

  return null;
}

/**
 * Check if a scope specification targets notes.
 */
export function isNoteScope(spec: ScopeSpec): boolean {
  if (typeof spec === 'string') {
    return parseNoteScope(spec) !== null;
  }
  if (typeof spec === 'function') {
    return false;
  }
  // Check if role is footnote or endnote
  if (spec.role) {
    const roles = Array.isArray(spec.role) ? spec.role : [spec.role];
    return roles.some(
      (r) => r === SemanticRole.Footnote || r === SemanticRole.Endnote
    );
  }
  return false;
}

// =============================================================================
// Scope Filter Creation
// =============================================================================

/**
 * Normalize any ScopeSpec into a ScopeFilter.
 *
 * @param spec - Any scope specification type
 * @returns Normalized ScopeFilter (empty if predicate)
 */
export function normalizeScope(spec: ScopeSpec): ScopeFilter {
  if (typeof spec === 'function') {
    return {}; // Can't normalize predicates
  }
  if (typeof spec === 'string') {
    return parseScope(spec);
  }
  return spec;
}

/**
 * Create a filter function from a ScopeSpec.
 *
 * Note: Section filtering requires tree context and is handled separately
 * in resolveScope(). This function handles all other filters.
 *
 * @param spec - Scope specification
 * @param options - Filter options
 * @returns Predicate function for filtering nodes
 */
export function createScopeFilter(
  spec: ScopeSpec,
  options: ScopeOptions = {}
): ScopePredicate {
  const { caseInsensitive = true } = options;

  // Handle predicate directly
  if (typeof spec === 'function') {
    return spec;
  }

  // Parse string to filter
  const filter: ScopeFilter = typeof spec === 'string' ? parseScope(spec) : spec;

  return (node: AccessibilityNode): boolean => {
    const nodeText = caseInsensitive
      ? (node.text ?? '').toLowerCase()
      : (node.text ?? '');

    // Contains filter
    if (filter.contains !== undefined) {
      const searchText = caseInsensitive
        ? filter.contains.toLowerCase()
        : filter.contains;
      if (!nodeText.includes(searchText)) return false;
    }

    // Not contains filter
    if (filter.notContains !== undefined) {
      const excludeText = caseInsensitive
        ? filter.notContains.toLowerCase()
        : filter.notContains;
      if (nodeText.includes(excludeText)) return false;
    }

    // Role filter
    if (filter.role !== undefined) {
      const roles = Array.isArray(filter.role) ? filter.role : [filter.role];
      if (!roles.includes(node.role)) return false;
    }

    // Style filter
    if (filter.style !== undefined) {
      const styles = Array.isArray(filter.style) ? filter.style : [filter.style];
      const nodeName = caseInsensitive
        ? node.style?.name?.toLowerCase()
        : node.style?.name;
      const styleMatches = styles.some((s) =>
        caseInsensitive ? s.toLowerCase() === nodeName : s === nodeName
      );
      if (!node.style?.name || !styleMatches) return false;
    }

    // Has changes filter
    if (filter.hasChanges !== undefined) {
      if ((node.hasChanges ?? false) !== filter.hasChanges) return false;
    }

    // Has comments filter
    if (filter.hasComments !== undefined) {
      if ((node.hasComments ?? false) !== filter.hasComments) return false;
    }

    // Level filter (for headings)
    if (filter.level !== undefined) {
      const levels = Array.isArray(filter.level) ? filter.level : [filter.level];
      if (node.level === undefined || !levels.includes(node.level)) return false;
    }

    // Length filters
    if (filter.minLength !== undefined) {
      if ((node.text?.length ?? 0) < filter.minLength) return false;
    }
    if (filter.maxLength !== undefined) {
      if ((node.text?.length ?? 0) > filter.maxLength) return false;
    }

    // Pattern filter (regex)
    if (filter.pattern !== undefined) {
      const flags = caseInsensitive ? 'i' : undefined;
      const regex = new RegExp(filter.pattern, flags);
      if (!regex.test(node.text ?? '')) return false;
    }

    // Refs filter (explicit ref list)
    if (filter.refs !== undefined && filter.refs.length > 0) {
      if (!filter.refs.includes(node.ref)) return false;
    }

    // Section filter is handled in resolveScope() with tree context

    return true;
  };
}

// =============================================================================
// Section Detection
// =============================================================================

/**
 * Flatten an accessibility tree to a list of nodes in document order.
 */
function flattenTree(tree: AccessibilityTree): AccessibilityNode[] {
  const nodes: AccessibilityNode[] = [];
  const content = tree.content ?? tree.outline ?? [];

  function flatten(node: AccessibilityNode): void {
    nodes.push(node);
    node.children?.forEach(flatten);
  }

  content.forEach(flatten);
  return nodes;
}

/**
 * Check if a node is a heading.
 */
function isHeadingNode(node: AccessibilityNode): boolean {
  // Check role
  if (node.role === SemanticRole.Heading) return true;

  // Check level (indicates heading)
  if (node.level !== undefined) return true;

  // Check style name
  const styleName = node.style?.name?.toLowerCase() ?? '';
  if (styleName.includes('heading')) return true;
  if (styleName === 'title') return true;

  return false;
}

/**
 * Find the section heading for a given node.
 *
 * Walks backwards through the flattened tree to find the first heading
 * that precedes this node.
 *
 * @param tree - The accessibility tree
 * @param nodeRef - Ref of the node
 * @returns The heading node, or undefined if not in a section
 */
export function findSectionHeading(
  tree: AccessibilityTree,
  nodeRef: Ref
): AccessibilityNode | undefined {
  const flatNodes = flattenTree(tree);

  // Find target node index
  const targetIndex = flatNodes.findIndex((n) => n.ref === nodeRef);
  if (targetIndex === -1) return undefined;

  // Walk backwards to find first heading
  for (let i = targetIndex - 1; i >= 0; i--) {
    const node = flatNodes[i];
    if (node && isHeadingNode(node)) {
      return node;
    }
  }

  return undefined;
}

/**
 * Check if a specific node is within a named section.
 *
 * Algorithm:
 * 1. Flatten tree to list of nodes in document order
 * 2. Find target node index
 * 3. Walk backwards from index-1
 * 4. Find first heading (by role or style)
 * 5. Check if heading text contains section name (case-insensitive)
 *
 * @param tree - The accessibility tree
 * @param nodeRef - Ref of the node to check
 * @param sectionName - Section name to match (case-insensitive)
 * @returns true if node is in the specified section
 */
export function isInSection(
  tree: AccessibilityTree,
  nodeRef: Ref,
  sectionName: string
): boolean {
  const heading = findSectionHeading(tree, nodeRef);
  if (!heading) return false;

  const headingText = (heading.text ?? '').toLowerCase();
  const searchName = sectionName.toLowerCase();

  return headingText.includes(searchName);
}

// =============================================================================
// Scope Resolution
// =============================================================================

/**
 * Resolve a scope specification against an accessibility tree.
 * Returns all matching nodes.
 *
 * @param tree - The accessibility tree to search
 * @param scope - Scope specification
 * @param options - Resolution options
 * @returns ScopeResult with matching nodes
 *
 * @example
 * // Find all paragraphs in "Methods" section
 * const result = resolveScope(tree, "section:Methods");
 *
 * // Find headings with tracked changes
 * const result = resolveScope(tree, { role: SemanticRole.Heading, hasChanges: true });
 */
export function resolveScope(
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options: ScopeOptions = {}
): ScopeResult {
  const { includeHeadings = false } = options;

  const filter = createScopeFilter(scope, options);
  const normalizedScope = normalizeScope(scope);

  const flatNodes = flattenTree(tree);
  const matches: AccessibilityNode[] = [];
  let totalEvaluated = 0;

  for (const node of flatNodes) {
    totalEvaluated++;

    // Check basic filter
    let passes = filter(node);

    // Handle section filter separately (needs tree context)
    if (passes && normalizedScope.section !== undefined) {
      passes = isInSection(tree, node.ref, normalizedScope.section);

      // Optionally exclude the heading itself
      if (passes && !includeHeadings && isHeadingNode(node)) {
        passes = false;
      }
    }

    if (passes) {
      matches.push(node);
    }
  }

  return {
    nodes: matches,
    totalEvaluated,
    scope: normalizedScope,
  };
}

/**
 * Filter nodes from a tree using a scope specification.
 * Convenience function that returns just the matching nodes.
 *
 * @param tree - The accessibility tree to search
 * @param scope - Scope specification
 * @param options - Resolution options
 * @returns Array of matching nodes
 */
export function filterByScope(
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options?: ScopeOptions
): AccessibilityNode[] {
  return resolveScope(tree, scope, options).nodes;
}

/**
 * Find the first node matching a scope specification.
 *
 * @param tree - The accessibility tree to search
 * @param scope - Scope specification
 * @param options - Resolution options
 * @returns First matching node, or undefined if none found
 */
export function findFirstByScope(
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options?: ScopeOptions
): AccessibilityNode | undefined {
  const result = resolveScope(tree, scope, options);
  return result.nodes[0];
}

/**
 * Count nodes matching a scope specification.
 *
 * @param tree - The accessibility tree to search
 * @param scope - Scope specification
 * @param options - Resolution options
 * @returns Number of matching nodes
 */
export function countByScope(
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options?: ScopeOptions
): number {
  return resolveScope(tree, scope, options).nodes.length;
}

/**
 * Get all refs matching a scope specification.
 *
 * @param tree - The accessibility tree to search
 * @param scope - Scope specification
 * @param options - Resolution options
 * @returns Array of refs for matching nodes
 */
export function getRefsByScope(
  tree: AccessibilityTree,
  scope: ScopeSpec,
  options?: ScopeOptions
): Ref[] {
  return resolveScope(tree, scope, options).nodes.map((n) => n.ref);
}
