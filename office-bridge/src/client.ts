import type { AppType, SessionInfo, ServerInfo } from './types.js';

// Re-export types for public API
export type { AppType } from './types.js';

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

export interface PageImageOptions {
  scale?: number;   // Default: 1.5 (higher = more detail)
  pages?: number[]; // Specific pages to get (default: all)
}

export interface PageImage {
  page: number;
  width: number;
  height: number;
  data: string;  // data:image/png;base64,...
}

export interface SlideImageOptions {
  width?: number;   // Target width in pixels
  height?: number;  // Target height in pixels (default: 600)
}

export interface SlideImage {
  slide: number;
  data: string;  // data:image/png;base64,...
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

  // =========================================================================
  // Page Image Methods
  // =========================================================================

  /**
   * Get all pages as images.
   * Exports document to PDF and converts each page to a PNG image.
   *
   * @param options - Scale and page selection options
   * @returns Array of page images with base64 data URLs
   *
   * @example
   * ```typescript
   * const pages = await doc.getPageImages();
   * console.log(`Document has ${pages.length} pages`);
   * // pages[0].data is "data:image/png;base64,..."
   * ```
   */
  getPageImages(options?: PageImageOptions): Promise<PageImage[]>;

  /**
   * Get a single page as an image.
   *
   * @param pageNum - Page number (1-indexed)
   * @param options - Scale options
   * @returns Page image or undefined if page doesn't exist
   *
   * @example
   * ```typescript
   * const page = await doc.getPageImage(1);
   * if (page) console.log('First page image:', page.data);
   * ```
   */
  getPageImage(pageNum: number, options?: Omit<PageImageOptions, 'pages'>): Promise<PageImage | undefined>;
}

// Base session interface for all apps (raw JS execution)
export interface OfficeSession {
  readonly id: string;
  readonly app: AppType;
  readonly filename: string;
  readonly path: string;
  readonly connectedAt: Date;
  readonly status: 'connected' | 'disconnected';

  /**
   * Execute raw JavaScript code in the app's context.
   */
  executeJs<T = unknown>(code: string, options?: ExecuteOptions): Promise<T>;
}

// Excel-specific session
export interface ExcelSession extends OfficeSession {
  readonly app: 'excel';
}

// PowerPoint-specific session
export interface PowerPointSession extends OfficeSession {
  readonly app: 'powerpoint';

  /**
   * Get all slides as images.
   *
   * @param options - Size options
   * @returns Array of slide images with base64 data URLs
   *
   * @example
   * ```typescript
   * const slides = await ppt.getSlideImages();
   * console.log(`Presentation has ${slides.length} slides`);
   * ```
   */
  getSlideImages(options?: SlideImageOptions): Promise<SlideImage[]>;

  /**
   * Get a single slide as an image.
   *
   * @param slideNum - Slide number (1-indexed)
   * @param options - Size options
   * @returns Slide image or undefined if slide doesn't exist
   *
   * @example
   * ```typescript
   * const slide = await ppt.getSlideImage(1);
   * if (slide) console.log('First slide image:', slide.data);
   * ```
   */
  getSlideImage(slideNum: number, options?: SlideImageOptions): Promise<SlideImage | undefined>;
}

// Outlook-specific session
export interface OutlookSession extends OfficeSession {
  readonly app: 'outlook';
}

export interface SessionFilter {
  app?: AppType;
}

export interface BridgeClient {
  /**
   * Get all connected sessions, optionally filtered by app.
   */
  sessions(filter?: SessionFilter): Promise<OfficeSession[]>;

  /**
   * Get connected Word documents (with full helper methods).
   */
  documents(): Promise<WordDocument[]>;

  /**
   * Get connected Excel workbooks.
   */
  excel(): Promise<ExcelSession[]>;

  /**
   * Get connected PowerPoint presentations.
   */
  powerpoint(): Promise<PowerPointSession[]>;

  /**
   * Get connected Outlook sessions.
   */
  outlook(): Promise<OutlookSession[]>;

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
  console.log(`Connected to Office Bridge (${serverInfo.documents} sessions)`);

  return {
    async sessions(filter?: SessionFilter): Promise<OfficeSession[]> {
      const url = filter?.app
        ? `${baseUrl}/sessions?app=${filter.app}`
        : `${baseUrl}/sessions`;
      const res = await fetch(url);
      if (!res.ok) {
        throw new Error(`Failed to list sessions: ${res.status}`);
      }
      const data = await res.json() as { sessions: SessionInfo[] };

      return data.sessions.map(info => createSessionHandle(baseUrl, info));
    },

    async documents(): Promise<WordDocument[]> {
      const res = await fetch(`${baseUrl}/sessions?app=word`);
      if (!res.ok) {
        throw new Error(`Failed to list documents: ${res.status}`);
      }
      const data = await res.json() as { sessions: SessionInfo[] };

      return data.sessions.map(info => createWordDocumentHandle(baseUrl, info));
    },

    async excel(): Promise<ExcelSession[]> {
      const res = await fetch(`${baseUrl}/sessions?app=excel`);
      if (!res.ok) {
        throw new Error(`Failed to list Excel workbooks: ${res.status}`);
      }
      const data = await res.json() as { sessions: SessionInfo[] };

      return data.sessions.map(info => createSessionHandle(baseUrl, info) as ExcelSession);
    },

    async powerpoint(): Promise<PowerPointSession[]> {
      const res = await fetch(`${baseUrl}/sessions?app=powerpoint`);
      if (!res.ok) {
        throw new Error(`Failed to list PowerPoint presentations: ${res.status}`);
      }
      const data = await res.json() as { sessions: SessionInfo[] };

      return data.sessions.map(info => createPowerPointSessionHandle(baseUrl, info));
    },

    async outlook(): Promise<OutlookSession[]> {
      const res = await fetch(`${baseUrl}/sessions?app=outlook`);
      if (!res.ok) {
        throw new Error(`Failed to list Outlook sessions: ${res.status}`);
      }
      const data = await res.json() as { sessions: SessionInfo[] };

      return data.sessions.map(info => createSessionHandle(baseUrl, info) as OutlookSession);
    },

    async close(): Promise<void> {
      // Nothing to clean up for HTTP client
      console.log('Disconnected from Office Bridge');
    },
  };
}

// Helper to create executeJs function for any session
function createExecuteJs(baseUrl: string, sessionId: string) {
  return async function executeJs<T = unknown>(code: string, options: ExecuteOptions = {}): Promise<T> {
    const timeout = options.timeout ?? 30000;

    const res = await fetch(`${baseUrl}/execute`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sessionId,
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
  };
}

// Generic session handle (works for any app)
function createSessionHandle(baseUrl: string, info: SessionInfo): OfficeSession {
  return {
    id: info.id,
    app: info.app,
    filename: info.filename,
    path: info.path,
    connectedAt: new Date(info.connectedAt),
    status: info.status,
    executeJs: createExecuteJs(baseUrl, info.id),
  };
}

// Word-specific document handle with full accessibility helpers
function createWordDocumentHandle(baseUrl: string, info: SessionInfo): WordDocument {
  const executeJs = createExecuteJs(baseUrl, info.id);

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

    // =========================================================================
    // Page Image Methods
    // =========================================================================

    async getPageImages(options?: PageImageOptions): Promise<PageImage[]> {
      const res = await fetch(`${baseUrl}/page-images`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId: info.id,
          scale: options?.scale,
          pages: options?.pages,
        }),
      });

      if (!res.ok) {
        const error = await res.text();
        throw new Error(`Failed to get page images: ${error}`);
      }

      const result = await res.json() as { success: boolean; images?: PageImage[]; error?: { message: string } };

      if (!result.success) {
        throw new Error(result.error?.message ?? 'Failed to get page images');
      }

      return result.images ?? [];
    },

    async getPageImage(pageNum: number, options?: Omit<PageImageOptions, 'pages'>): Promise<PageImage | undefined> {
      const pages = await this.getPageImages({ ...options, pages: [pageNum] });
      return pages.find(p => p.page === pageNum);
    },
  };
}

function createPowerPointSessionHandle(baseUrl: string, info: SessionInfo): PowerPointSession {
  const executeJs = createExecuteJs(baseUrl, info.id);

  return {
    id: info.id,
    app: 'powerpoint',
    filename: info.filename,
    path: info.path,
    connectedAt: new Date(info.connectedAt),
    status: info.status,

    executeJs,

    // =========================================================================
    // Slide Image Methods
    // =========================================================================

    async getSlideImages(options?: SlideImageOptions): Promise<SlideImage[]> {
      const height = options?.height ?? 600;
      const width = options?.width;

      const optionsJson = JSON.stringify({ height, width });
      const code = `
        const options = ${optionsJson};
        const presentation = context.presentation;
        const slideCount = presentation.slides.getCount();
        await context.sync();

        const images = [];
        for (let i = 0; i < slideCount.value; i++) {
          const slide = presentation.slides.getItemAt(i);
          const imageOpts = {};
          if (options.height) imageOpts.height = options.height;
          if (options.width) imageOpts.width = options.width;
          const imageResult = slide.getImageAsBase64(imageOpts);
          await context.sync();
          images.push({
            slide: i + 1,
            data: 'data:image/png;base64,' + imageResult.value
          });
        }
        return images;
      `;
      return executeJs<SlideImage[]>(code);
    },

    async getSlideImage(slideNum: number, options?: SlideImageOptions): Promise<SlideImage | undefined> {
      const height = options?.height ?? 600;
      const width = options?.width;

      const optionsJson = JSON.stringify({ height, width, slideNum });
      const code = `
        const options = ${optionsJson};
        const presentation = context.presentation;
        const slideCount = presentation.slides.getCount();
        await context.sync();

        if (options.slideNum < 1 || options.slideNum > slideCount.value) {
          return undefined;
        }

        const slide = presentation.slides.getItemAt(options.slideNum - 1);
        const imageOpts = {};
        if (options.height) imageOpts.height = options.height;
        if (options.width) imageOpts.width = options.width;
        const imageResult = slide.getImageAsBase64(imageOpts);
        await context.sync();

        return {
          slide: options.slideNum,
          data: 'data:image/png;base64,' + imageResult.value
        };
      `;
      return executeJs<SlideImage | undefined>(code);
    },
  };
}
