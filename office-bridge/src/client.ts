import type { DocumentInfo, ServerInfo } from './types.js';

// Re-export accessibility types for public API
export type {
  AccessibilityTree,
  AccessibilityNode,
  TreeOptions,
  VerbosityLevel,
  ChangeViewMode,
  ContentMode,
  ViewMode,
  Ref,
  SemanticRole,
  TrackedChange,
  Comment,
  EditResult,
} from './accessibility/index.js';

export type { EditOptions, FormatOptions } from './accessibility/editing.js';

export interface ConnectOptions {
  port?: number;  // Default: 3847
  host?: string;  // Default: 'localhost'
}

export interface ExecuteOptions {
  timeout?: number;  // Default: 30000 (30 seconds)
}

export interface ConsoleEntry {
  level: 'log' | 'warn' | 'error' | 'info';
  message: string;
  timestamp: Date;
  documentId: string;
}

// Import types for method signatures
import type {
  AccessibilityTree,
  TreeOptions,
  Ref,
  EditResult,
} from './accessibility/index.js';
import type { EditOptions, FormatOptions } from './accessibility/editing.js';

export interface WordDocument {
  readonly id: string;
  readonly filename: string;
  readonly path: string;
  readonly connectedAt: Date;
  readonly status: 'connected' | 'disconnected';

  /**
   * Execute raw JavaScript code in Word.run() context.
   * Low-level API for advanced use cases.
   */
  executeJs<T = unknown>(code: string, options?: ExecuteOptions): Promise<T>;

  // =========================================================================
  // Accessibility Tree Methods
  // =========================================================================

  /**
   * Get accessibility tree as YAML string.
   *
   * @param options - Tree building options (verbosity, view mode, etc.)
   * @returns YAML string representation of the document tree
   *
   * @example
   * ```typescript
   * // Get minimal outline for navigation
   * const outline = await doc.getTree({ verbosity: 'minimal' });
   *
   * // Get standard content view
   * const content = await doc.getTree({ verbosity: 'standard' });
   *
   * // Get full detail with run-level formatting
   * const full = await doc.getTree({ verbosity: 'full' });
   * ```
   */
  getTree(options?: TreeOptions): Promise<string>;

  /**
   * Get raw accessibility tree object.
   *
   * @param options - Tree building options (verbosity, view mode, etc.)
   * @returns AccessibilityTree object
   *
   * @example
   * ```typescript
   * const tree = await doc.getTreeRaw();
   * console.log(tree.document.stats.paragraphs);
   * ```
   */
  getTreeRaw(options?: TreeOptions): Promise<AccessibilityTree>;

  // =========================================================================
  // Ref-Based Editing Methods
  // =========================================================================

  /**
   * Replace text content at a ref.
   *
   * @param ref - Reference to the element to replace
   * @param newText - New text to replace with
   * @param options - Edit options (track changes, author, etc.)
   * @returns EditResult indicating success/failure
   *
   * @example
   * ```typescript
   * const result = await doc.replaceByRef('p:3', 'Updated paragraph text', { track: true });
   * if (result.success) console.log('Replaced:', result.newRef);
   * ```
   */
  replaceByRef(ref: Ref, newText: string, options?: EditOptions): Promise<EditResult>;

  /**
   * Insert content after a ref.
   *
   * @param ref - Reference to insert after
   * @param content - Content to insert
   * @param options - Edit options (track changes, author, etc.)
   * @returns EditResult indicating success/failure
   *
   * @example
   * ```typescript
   * await doc.insertAfterRef('p:5', ' (amended)', { track: true });
   * ```
   */
  insertAfterRef(ref: Ref, content: string, options?: EditOptions): Promise<EditResult>;

  /**
   * Insert content before a ref.
   *
   * @param ref - Reference to insert before
   * @param content - Content to insert
   * @param options - Edit options (track changes, author, etc.)
   * @returns EditResult indicating success/failure
   *
   * @example
   * ```typescript
   * await doc.insertBeforeRef('p:5', 'Note: ', { track: true });
   * ```
   */
  insertBeforeRef(ref: Ref, content: string, options?: EditOptions): Promise<EditResult>;

  /**
   * Delete element at a ref.
   *
   * @param ref - Reference to delete
   * @param options - Edit options (track changes, author, etc.)
   * @returns EditResult indicating success/failure
   *
   * @example
   * ```typescript
   * await doc.deleteByRef('p:3', { track: true });
   * ```
   */
  deleteByRef(ref: Ref, options?: EditOptions): Promise<EditResult>;

  /**
   * Apply formatting at a ref.
   *
   * @param ref - Reference to format
   * @param formatting - Formatting options to apply
   * @returns EditResult indicating success/failure
   *
   * @example
   * ```typescript
   * await doc.formatByRef('p:3', { bold: true, color: '#0000FF' });
   * ```
   */
  formatByRef(ref: Ref, formatting: FormatOptions): Promise<EditResult>;

  /**
   * Get text content at a ref.
   *
   * @param ref - Reference to read from
   * @returns Text content or undefined if not found
   *
   * @example
   * ```typescript
   * const text = await doc.getTextByRef('p:3');
   * console.log('Paragraph text:', text);
   * ```
   */
  getTextByRef(ref: Ref): Promise<string | undefined>;
}

export interface BridgeClient {
  documents(): Promise<WordDocument[]>;
  close(): Promise<void>;
}

export async function connect(options: ConnectOptions = {}): Promise<BridgeClient> {
  const host = options.host ?? 'localhost';
  const port = options.port ?? 3847;
  const baseUrl = `https://${host}:${port}`;

  // Fetch server info to verify connection
  const infoRes = await fetch(baseUrl);
  if (!infoRes.ok) {
    throw new Error(`Failed to connect to bridge server: ${infoRes.status}`);
  }
  const serverInfo = await infoRes.json() as ServerInfo;
  console.log(`Connected to Office Bridge (${serverInfo.documents} documents)`);

  return {
    async documents(): Promise<WordDocument[]> {
      const res = await fetch(`${baseUrl}/documents`);
      if (!res.ok) {
        throw new Error(`Failed to list documents: ${res.status}`);
      }
      const data = await res.json() as { documents: DocumentInfo[] };

      return data.documents.map(doc => createDocumentHandle(baseUrl, doc));
    },

    async close(): Promise<void> {
      // Nothing to clean up for HTTP client
      console.log('Disconnected from Office Bridge');
    },
  };
}

function createDocumentHandle(baseUrl: string, info: DocumentInfo): WordDocument {
  // Helper function to execute JS code via the bridge
  async function executeJs<T = unknown>(code: string, options: ExecuteOptions = {}): Promise<T> {
    const timeout = options.timeout ?? 30000;

    const res = await fetch(`${baseUrl}/execute`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        documentId: info.id,
        code,
        timeout,
      }),
    });

    if (!res.ok) {
      const error = await res.text();
      throw new Error(`Execution failed: ${error}`);
    }

    const result = await res.json() as { success: boolean; result?: T; error?: { message: string } };

    if (!result.success) {
      throw new Error(result.error?.message ?? 'Execution failed');
    }

    return result.result as T;
  }

  return {
    id: info.id,
    filename: info.filename,
    path: info.path,
    connectedAt: new Date(info.connectedAt),
    status: info.status,

    executeJs,

    // =========================================================================
    // Accessibility Tree Methods
    // =========================================================================

    async getTree(options?: TreeOptions): Promise<string> {
      const optionsJson = JSON.stringify(options ?? {});
      // Note: Requires DocTree global to be exposed by the add-in
      // The add-in must bundle accessibility modules and expose them as window.DocTree
      const code = `
        const options = ${optionsJson};
        const tree = await DocTree.buildTree(context, options);
        return DocTree.treeToYaml(tree, options.verbosity ?? 'standard');
      `;
      return executeJs<string>(code);
    },

    async getTreeRaw(options?: TreeOptions): Promise<AccessibilityTree> {
      const optionsJson = JSON.stringify(options ?? {});
      const code = `
        const options = ${optionsJson};
        return await DocTree.buildTree(context, options);
      `;
      return executeJs<AccessibilityTree>(code);
    },

    // =========================================================================
    // Ref-Based Editing Methods
    // =========================================================================

    async replaceByRef(ref: Ref, newText: string, options?: EditOptions): Promise<EditResult> {
      const refJson = JSON.stringify(ref);
      const newTextJson = JSON.stringify(newText);
      const optionsJson = JSON.stringify(options ?? {});
      const code = `
        return await DocTree.replaceByRef(context, ${refJson}, ${newTextJson}, ${optionsJson});
      `;
      return executeJs<EditResult>(code);
    },

    async insertAfterRef(ref: Ref, content: string, options?: EditOptions): Promise<EditResult> {
      const refJson = JSON.stringify(ref);
      const contentJson = JSON.stringify(content);
      const optionsJson = JSON.stringify(options ?? {});
      const code = `
        return await DocTree.insertAfterRef(context, ${refJson}, ${contentJson}, ${optionsJson});
      `;
      return executeJs<EditResult>(code);
    },

    async insertBeforeRef(ref: Ref, content: string, options?: EditOptions): Promise<EditResult> {
      const refJson = JSON.stringify(ref);
      const contentJson = JSON.stringify(content);
      const optionsJson = JSON.stringify(options ?? {});
      const code = `
        return await DocTree.insertBeforeRef(context, ${refJson}, ${contentJson}, ${optionsJson});
      `;
      return executeJs<EditResult>(code);
    },

    async deleteByRef(ref: Ref, options?: EditOptions): Promise<EditResult> {
      const refJson = JSON.stringify(ref);
      const optionsJson = JSON.stringify(options ?? {});
      const code = `
        return await DocTree.deleteByRef(context, ${refJson}, ${optionsJson});
      `;
      return executeJs<EditResult>(code);
    },

    async formatByRef(ref: Ref, formatting: FormatOptions): Promise<EditResult> {
      const refJson = JSON.stringify(ref);
      const formattingJson = JSON.stringify(formatting);
      const code = `
        return await DocTree.formatByRef(context, ${refJson}, ${formattingJson});
      `;
      return executeJs<EditResult>(code);
    },

    async getTextByRef(ref: Ref): Promise<string | undefined> {
      const refJson = JSON.stringify(ref);
      const code = `
        return await DocTree.getTextByRef(context, ${refJson});
      `;
      return executeJs<string | undefined>(code);
    },
  };
}
