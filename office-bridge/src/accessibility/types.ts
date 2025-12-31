/**
 * DocTree Accessibility Layer Types
 *
 * TypeScript types for the DOCX accessibility tree structure.
 * Based on docs/DOCTREE_SPEC.md - provides a semantic, agent-friendly
 * representation of Word documents with stable refs for precise editing.
 */

// =============================================================================
// Ref System (Section 2 of spec)
// =============================================================================

/**
 * Reference string for addressing document elements.
 *
 * Format: element_type ":" identifier ["/" sub_element]*
 *
 * Examples:
 *   p:3                     - 4th paragraph (0-indexed)
 *   p:~xK4mNp2q             - Paragraph by fingerprint (stable)
 *   tbl:0/row:2/cell:1      - Table 0, row 2, cell 1
 *   tbl:0/row:2/cell:1/p:0  - First paragraph in that cell
 *   ins:42                  - Tracked insertion with id=42
 *   hdr:default/p:0         - First paragraph in default header
 */
export type Ref = string;

/**
 * Element type prefixes used in refs.
 */
export type RefPrefix =
  | 'p'      // Paragraph
  | 'r'      // Run (text span)
  | 'tbl'    // Table
  | 'row'    // Table row
  | 'cell'   // Table cell
  | 'ins'    // Tracked insertion
  | 'del'    // Tracked deletion
  | 'hdr'    // Header
  | 'ftr'    // Footer
  | 'fn'     // Footnote
  | 'en'     // Endnote
  | 'cmt'    // Comment
  | 'sec'    // Section
  | 'img'    // Image
  | 'chart'  // Chart
  | 'diagram'// SmartArt diagram
  | 'shape'  // Shape
  | 'obj'    // OLE object
  | 'vml'    // Legacy VML
  | 'bk'     // Bookmark
  | 'lnk'    // Hyperlink
  | 'xref'   // Cross-reference
  | 'toc'    // Table of Contents entry
  | 'change' // Tracked change reference
;

// =============================================================================
// Semantic Roles (Section 3 of spec)
// =============================================================================

/**
 * Semantic roles for document elements.
 * Based on ARIA-like taxonomy adapted for Word documents.
 */
export enum SemanticRole {
  // Document Roles (Landmarks)
  Document = 'document',
  Header = 'header',
  Footer = 'footer',
  Section = 'section',

  // Structural Roles
  Heading = 'heading',
  Paragraph = 'paragraph',
  Blockquote = 'blockquote',
  List = 'list',
  ListItem = 'listitem',
  Table = 'table',
  Row = 'row',
  Cell = 'cell',

  // Inline Roles
  Text = 'text',
  Strong = 'strong',
  Emphasis = 'emphasis',
  Link = 'link',

  // Annotation Roles
  Insertion = 'insertion',
  Deletion = 'deletion',
  Comment = 'comment',

  // Image and Object Roles
  Image = 'image',
  Chart = 'chart',
  Diagram = 'diagram',
  Shape = 'shape',

  // Reference Roles
  Footnote = 'footnote',
  Endnote = 'endnote',
  Bookmark = 'bookmark',
}

// =============================================================================
// View Modes (Sections 4.1 and 5.1 of spec)
// =============================================================================

/**
 * Verbosity levels for tree output.
 *
 * - minimal: Structure overview, navigation (~2000 tokens)
 * - standard: Full content, edit planning (default)
 * - full: Complete fidelity, run-level detail
 */
export type VerbosityLevel = 'minimal' | 'standard' | 'full';

/**
 * View modes for tracked changes display.
 *
 * - final: Show as if changes accepted
 * - original: Show as if changes rejected
 * - markup: Show all changes with markers
 */
export type ChangeViewMode = 'final' | 'original' | 'markup';

/**
 * Content mode for editing focus.
 *
 * - content: Text and structure (lower token cost)
 * - styling: Includes run-level formatting (higher token cost)
 */
export type ContentMode = 'content' | 'styling';

/**
 * Combined view mode for tree generation.
 */
export interface ViewMode {
  /** Include document body content */
  includeBody?: boolean;
  /** Include headers and footers */
  includeHeaders?: boolean;
  /** Include comments */
  includeComments?: boolean;
  /** Include tracked changes */
  includeTrackedChanges?: boolean;
  /** Include formatting details */
  includeFormatting?: boolean;
}

// =============================================================================
// Tree Options (Section 6.2 of spec)
// =============================================================================

/**
 * Configuration for section detection heuristics.
 */
export interface SectionDetectionConfig {
  /** Detect bold first-line as heading */
  detectBoldFirstLine?: boolean;
  /** Detect ALL CAPS first-line as heading */
  detectCapsFirstLine?: boolean;
  /** Detect numbered sections (1., 2., Article I) */
  detectNumberedSections?: boolean;
  /** Detect blank line breaks as section dividers */
  detectBlankLineBreaks?: boolean;
  /** Minimum paragraphs per section */
  minSectionParagraphs?: number;
  /** Maximum heading length in characters */
  maxHeadingLength?: number;
  /** Custom numbering patterns (regex) */
  numberingPatterns?: string[];
}

/**
 * Options for building accessibility trees.
 */
export interface TreeOptions {
  /** Verbosity level for output */
  verbosity?: VerbosityLevel;
  /** Change view mode */
  changeViewMode?: ChangeViewMode;
  /** Content vs styling mode */
  contentMode?: ContentMode;
  /** View mode options */
  viewMode?: ViewMode;
  /** Maximum tokens for output (triggers truncation) */
  maxTokens?: number;
  /** Section detection configuration */
  sectionDetection?: SectionDetectionConfig;
  /** Scope to specific refs only (legacy, use scopeFilter instead) */
  scopeRefs?: Ref[];
  /** Filter tree by scope specification */
  scopeFilter?: ScopeSpec;
}

// =============================================================================
// Tracked Changes (Section 5 of spec)
// =============================================================================

/**
 * Type of tracked change.
 */
export type TrackedChangeType = 'ins' | 'del';

/**
 * Tracked change information.
 */
export interface TrackedChange {
  /** Reference to this change */
  ref: Ref;
  /** Type of change (insertion or deletion) */
  type: TrackedChangeType;
  /** Word revision ID */
  id?: string;
  /** Author of the change */
  author: string;
  /** Date of the change (ISO 8601) */
  date: string;
  /** Text content of the change */
  text: string;
  /** Location in document */
  location?: {
    paragraphRef: Ref;
    runRef?: Ref;
  };
}

/**
 * Format change information (tracked formatting changes).
 */
export interface FormatChange {
  /** Author of the format change */
  author: string;
  /** Date of the change (ISO 8601) */
  date: string;
  /** Properties before change */
  before: Record<string, unknown>;
  /** Properties after change */
  after: Record<string, unknown>;
}

// =============================================================================
// Comments (Section 5.4 of spec)
// =============================================================================

/**
 * Comment reply in a thread.
 */
export interface CommentReply {
  /** Reference to this reply */
  ref: Ref;
  /** Author of the reply */
  author: string;
  /** Date of the reply (ISO 8601) */
  date: string;
  /** Reply text content */
  text: string;
}

/**
 * Comment annotation on document content.
 */
export interface Comment {
  /** Reference to this comment */
  ref: Ref;
  /** Comment ID from Word */
  id: string;
  /** Author of the comment */
  author: string;
  /** Date of the comment (ISO 8601) */
  date: string;
  /** Comment text content */
  text: string;
  /** Text the comment is attached to */
  onText?: string;
  /** Whether the comment is resolved */
  resolved: boolean;
  /** Reply thread */
  replies?: CommentReply[];
}

// =============================================================================
// Accessibility Node (Section 6.1 of spec)
// =============================================================================

/**
 * Style information for a node.
 */
export interface NodeStyle {
  /** Named style (e.g., "Heading1", "Normal") */
  name?: string;
  /** Direct formatting overrides */
  formatting?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    font?: string;
    size?: string;
    color?: string;
    highlight?: string;
  };
  /** Paragraph formatting (content mode) */
  paragraph?: {
    alignment?: 'left' | 'center' | 'right' | 'justify';
    spacingBefore?: string;
    spacingAfter?: string;
    lineSpacing?: string;
    indent?: string;
  };
}

/**
 * Image information.
 */
export interface ImageInfo {
  /** Reference to the image */
  ref: Ref;
  /** Name/title of the image */
  name?: string;
  /** Alt text for accessibility */
  altText?: string;
  /** Size in human-readable format */
  size?: string;
  /** Detailed size in EMUs (English Metric Units) */
  sizeEmu?: {
    widthEmu: number;
    heightEmu: number;
  };
  /** Image format (png, jpg, etc.) */
  format?: string;
  /** Relationship ID in OOXML */
  relationshipId?: string;
  /** Position type */
  positionType?: 'inline' | 'floating';
  /** Floating position details */
  floatingPosition?: {
    horizontal: string;
    vertical: string;
    relativeTo: string;
    wrap?: string;
  };
}

/**
 * Link information.
 */
export interface LinkInfo {
  /** Reference to the link */
  ref: Ref;
  /** Displayed link text */
  text: string;
  /** Internal target (bookmark ref) */
  target?: Ref;
  /** Target location ref */
  targetLocation?: Ref;
  /** External URL */
  url?: string;
}

/**
 * Bookmark information.
 */
export interface BookmarkInfo {
  /** Reference to the bookmark */
  ref: Ref;
  /** Bookmark name */
  name: string;
  /** Location in document */
  location: Ref;
  /** Preview of bookmarked text */
  textPreview?: string;
  /** Refs that reference this bookmark */
  referencedBy?: Ref[];
}

/**
 * Footnote information.
 */
export interface FootnoteInfo {
  /** Reference to the footnote (fn:1, fn:2, etc.) */
  ref: Ref;
  /** Footnote ID from Word */
  id: number;
  /** Footnote text content */
  text: string;
  /** Paragraph ref where this footnote is referenced */
  referencedFrom?: Ref;
}

/**
 * Endnote information.
 */
export interface EndnoteInfo {
  /** Reference to the endnote (en:1, en:2, etc.) */
  ref: Ref;
  /** Endnote ID from Word */
  id: number;
  /** Endnote text content */
  text: string;
  /** Paragraph ref where this endnote is referenced */
  referencedFrom?: Ref;
}

/**
 * Node representing an element in the accessibility tree.
 */
export interface AccessibilityNode {
  /** Stable reference to this node */
  ref: Ref;
  /** Semantic role of this node */
  role: SemanticRole;
  /** Text content (may include change markers in markup mode) */
  text?: string;
  /** Child nodes */
  children?: AccessibilityNode[];
  /** Style information */
  style?: NodeStyle;
  /** Tracked changes affecting this node */
  changes?: TrackedChange[];
  /** Change refs for quick lookup */
  changeRefs?: Ref[];
  /** Whether this node has tracked changes */
  hasChanges?: boolean;
  /** Comments on this node */
  comments?: Comment[];
  /** Comment refs for quick lookup */
  commentRefs?: Ref[];
  /** Whether this node has comments */
  hasComments?: boolean;
  /** Images in this node */
  images?: ImageInfo[];
  /** Floating images anchored at this node */
  floatingImages?: ImageInfo[];
  /** Links in this node */
  links?: LinkInfo[];
  /** Bookmark at this node */
  bookmark?: BookmarkInfo;
  /** Incoming references to this node */
  incomingReferences?: Ref[];
  /** Footnote refs referenced in this node */
  footnoteRefs?: Ref[];
  /** Endnote refs referenced in this node */
  endnoteRefs?: Ref[];

  // Role-specific properties
  /** Heading level (for heading role) */
  level?: number;
  /** Whether this is a header row (for row role) */
  isHeader?: boolean;
  /** Table dimensions (for table role) */
  dimensions?: { rows: number; cols: number };
  /** Run-level information (for full verbosity) */
  runs?: AccessibilityNode[];
  /** Format change (for tracked format changes) */
  formatChange?: FormatChange;

  // Section-specific (for outline mode)
  /** Heading ref for section */
  headingRef?: Ref;
  /** Paragraph count in section */
  paragraphCount?: number;
  /** Table refs in section */
  tables?: Ref[];
  /** Tracked change count in section */
  trackedChangesCount?: number;
  /** Preview text for section */
  preview?: string;
  /** Section detection method */
  detection?: 'heading_style' | 'outline_level' | 'bold_heuristic' | 'caps_heuristic' | 'numbered_section' | 'fallback';
  /** Detection confidence */
  confidence?: 'high' | 'medium' | 'low';
}

// =============================================================================
// Tree Structure (Section 4.3 of spec)
// =============================================================================

/**
 * Document statistics.
 */
export interface DocumentStats {
  /** Total paragraph count */
  paragraphs: number;
  /** Total table count */
  tables: number;
  /** Total tracked change count */
  trackedChanges: number;
  /** Total comment count */
  comments: number;
  /** Total footnote count */
  footnotes?: number;
  /** Total endnote count */
  endnotes?: number;
  /** Total section count (for outline mode) */
  sections?: number;
}

/**
 * Section detection metadata.
 */
export interface SectionDetectionInfo {
  /** Detection method used */
  method: 'heading_style' | 'heuristic' | 'fallback';
  /** Overall confidence */
  confidence: 'high' | 'medium' | 'low';
}

/**
 * Document metadata for tree header.
 */
export interface DocumentMetadata {
  /** Document path or filename */
  path?: string;
  /** Verbosity level of this tree */
  verbosity: VerbosityLevel;
  /** Document statistics */
  stats: DocumentStats;
  /** Section detection info (outline mode) */
  sectionDetection?: SectionDetectionInfo;
  /** Current mode */
  mode?: 'outline' | 'content' | 'styling';
}

/**
 * Link summary for document.
 */
export interface LinkSummary {
  /** Internal links */
  internal: LinkInfo[];
  /** External links */
  external: LinkInfo[];
  /** Broken links */
  broken: Array<LinkInfo & { error: string }>;
}

/**
 * Complete accessibility tree for a document.
 */
export interface AccessibilityTree {
  /** Document metadata */
  document: DocumentMetadata;
  /** Content nodes (standard/full mode) */
  content?: AccessibilityNode[];
  /** Outline nodes (outline mode) */
  outline?: AccessibilityNode[];
  /** All tracked changes */
  trackedChanges?: TrackedChange[];
  /** All comments */
  comments?: Comment[];
  /** All bookmarks */
  bookmarks?: BookmarkInfo[];
  /** All footnotes */
  footnotes?: FootnoteInfo[];
  /** All endnotes */
  endnotes?: EndnoteInfo[];
  /** Link summary */
  links?: LinkSummary;
  /** Navigation hints (outline mode) */
  navigation?: {
    expandSection: string;
    search: string;
  };
}

// =============================================================================
// Error Types (Section 6.4 of spec)
// =============================================================================

/**
 * Error thrown when a ref cannot be resolved.
 */
export interface RefNotFoundError {
  type: 'RefNotFoundError';
  ref: Ref;
  message: string;
}

/**
 * Error thrown when a ref points to a deleted element.
 */
export interface StaleRefError {
  type: 'StaleRefError';
  ref: Ref;
  message: string;
}

/**
 * Union of accessibility-related errors.
 */
export type AccessibilityError = RefNotFoundError | StaleRefError;

// =============================================================================
// Edit Types (Section 6.3 of spec)
// =============================================================================

/**
 * Position for insertions.
 */
export type InsertPosition = 'before' | 'after' | 'start' | 'end';

/**
 * Result of an edit operation.
 */
export interface EditResult {
  /** Whether the edit succeeded */
  success: boolean;
  /** New ref for inserted/modified element */
  newRef?: Ref;
  /** Error message if failed */
  error?: string;
}

// =============================================================================
// Scope System (Section 7 of spec)
// =============================================================================

/**
 * Dictionary-based filter with AND logic.
 * All specified criteria must match for a node to be included.
 *
 * @example
 * // Find paragraphs in "Methods" section containing "results"
 * { section: "Methods", contains: "results" }
 *
 * // Find headings with tracked changes
 * { role: SemanticRole.Heading, hasChanges: true }
 */
export interface ScopeFilter {
  /** Text that must be contained in the node */
  contains?: string;
  /** Text that must NOT be in the node */
  notContains?: string;
  /** Section heading name (node must be under this heading) */
  section?: string;
  /** Semantic role(s) to match */
  role?: SemanticRole | SemanticRole[];
  /** Style name(s) to match */
  style?: string | string[];
  /** Only nodes with tracked changes */
  hasChanges?: boolean;
  /** Only nodes with comments */
  hasComments?: boolean;
  /** Heading level(s) to match (for heading role) */
  level?: number | number[];
  /** Regex pattern to match against text */
  pattern?: string;
  /** Specific refs to include */
  refs?: Ref[];
  /** Minimum text length */
  minLength?: number;
  /** Maximum text length */
  maxLength?: number;
}

/**
 * Custom predicate function for flexible node filtering.
 *
 * @example
 * // Custom filter for paragraphs starting with "WHEREAS"
 * (node) => node.text?.startsWith("WHEREAS")
 */
export type ScopePredicate = (node: AccessibilityNode) => boolean;

/**
 * Union type for all scope specification formats.
 *
 * Supports:
 * - String shortcuts: "keyword", "section:Name", "role:heading"
 * - Dictionary filters: { contains: "text", section: "Intro" }
 * - Predicate functions: (node) => boolean
 *
 * @example
 * // String shortcut
 * "section:Introduction"
 *
 * // Dictionary filter
 * { contains: "payment", notContains: "Exhibit" }
 *
 * // Predicate function
 * (node) => node.text?.length > 100
 */
export type ScopeSpec = string | ScopeFilter | ScopePredicate;

/**
 * Options for scope resolution.
 */
export interface ScopeOptions {
  /** Whether to include heading nodes when scoping by section (default: false) */
  includeHeadings?: boolean;
  /** Case-insensitive text matching (default: true) */
  caseInsensitive?: boolean;
}

/**
 * Result of scope resolution with metadata.
 */
export interface ScopeResult {
  /** Matching nodes */
  nodes: AccessibilityNode[];
  /** Total nodes evaluated */
  totalEvaluated: number;
  /** Normalized scope filter used */
  scope: ScopeFilter;
}

/**
 * Result of parsing a note-related scope (footnotes/endnotes).
 */
export interface NoteScope {
  /** Type of note scope */
  scopeType: 'footnotes' | 'endnotes' | 'notes' | 'footnote' | 'endnote';
  /** Specific note ID (for 'footnote:N' or 'endnote:N' scopes) */
  noteId?: string;
}
