/// <reference types="vite/client" />

type MenuAction =
  | 'new-file'
  | 'open-file'
  | 'import-docx'
  | 'save-file'
  | 'save-file-as'
  | 'reference-image'
  | 'open-ai-chat'
  | 'open-ai-config'
  | 'ai-rewrite'
  | 'ai-expand'
  | 'ai-shorten'
  | 'ai-continue'
  | 'ai-beautify'
  | 'ai-translate'
  | 'export-html'
  | 'export-pdf'
  | 'theme-paper'
  | 'theme-mist'
  | 'theme-slate'
  | 'theme-graphite'
  | 'theme-terminal'
  | 'theme-nightcode'
  | 'theme-campus'
  | 'theme-youth'
  | 'view-write'
  | 'view-split'
  | 'view-preview';

type OpenFileResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      filePath?: string;
      content?: string;
    };

type DocxImportImageAsset = {
  token: string;
  originalName: string;
  contentType: string;
  dataUrl: string;
};

type DocxImportSummary = {
  headingCount: number;
  listCount: number;
  tableCount: number;
  imageCount: number;
  codeBlockCount: number;
};

type OpenDocxFileResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      sourceFilePath: string;
      sourceFileName: string;
      markdown: string;
      previewHtml: string;
      messages: string[];
      summary: DocxImportSummary;
      imageAssets: DocxImportImageAsset[];
    };

type SaveFileResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      filePath?: string;
    };

type SaveFilePayload = {
  filePath: string | null;
  content: string;
  forceDialog?: boolean;
};

type SyncDocumentStatePayload = {
  filePath: string | null;
  displayName: string;
  content: string;
  isDirty: boolean;
};

type SaveImportedMarkdownPayload = {
  defaultPath: string;
  markdown: string;
  imageAssets: DocxImportImageAsset[];
};

type SaveImportedMarkdownResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      filePath: string;
      content: string;
    };

type ReadFileResult = {
  filePath: string;
  content: string;
};

type MaterializeImportedMarkdownPayload = {
  documentPath: string;
  markdown: string;
  imageAssets: DocxImportImageAsset[];
};

type MaterializeImportedMarkdownResult = {
  canceled: false;
  content: string;
};

type ExportHtmlPayload = {
  defaultPath: string;
  html: string;
};

type ExportPdfPayload = {
  defaultPath: string;
};

type SaveImageAssetPayload = {
  documentPath: string | null;
  originalName: string;
  dataUrl: string;
};

type PickImageReferencePayload = {
  documentPath: string | null;
};

type PickImageReferenceResult =
  | {
      canceled: true;
    }
  | {
      canceled: false;
      filePath: string;
      markdownPath: string;
      displayName: string;
    };

type PersistedAIProfile = {
  name: string;
  provider: 'builtin' | 'openai-compatible';
  apiBase: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxTokens: number;
};

type ReadAiConfigDocumentResult = {
  filePath: string;
  activeProfileName: string | null;
  profiles: PersistedAIProfile[];
};

type WriteAiConfigDocumentPayload = {
  activeProfileName: string | null;
  profiles: PersistedAIProfile[];
};

type WriteAiConfigDocumentResult = {
  filePath: string;
  activeProfileName: string | null;
  profiles: PersistedAIProfile[];
};

interface Window {
  markdownApp?: {
    ping: () => string;
    openMarkdownFile: () => Promise<OpenFileResult>;
    openDocxFile: () => Promise<OpenDocxFileResult>;
    saveMarkdownFile: (payload: SaveFilePayload) => Promise<SaveFileResult>;
    saveImportedMarkdownFile: (payload: SaveImportedMarkdownPayload) => Promise<SaveImportedMarkdownResult>;
    materializeImportedMarkdown: (
      payload: MaterializeImportedMarkdownPayload
    ) => Promise<MaterializeImportedMarkdownResult>;
    readMarkdownFile: (filePath: string) => Promise<ReadFileResult>;
    exportHtmlFile: (payload: ExportHtmlPayload) => Promise<SaveFileResult>;
    exportPdfFile: (payload: ExportPdfPayload) => Promise<SaveFileResult>;
    saveImageAsset: (payload: SaveImageAssetPayload) => Promise<SaveFileResult>;
    pickImageReference: (payload: PickImageReferencePayload) => Promise<PickImageReferenceResult>;
    readAiConfigDocument: () => Promise<ReadAiConfigDocumentResult>;
    writeAiConfigDocument: (payload: WriteAiConfigDocumentPayload) => Promise<WriteAiConfigDocumentResult>;
    openExternalLink: (url: string) => Promise<{ success: boolean }>;
    syncDocumentState: (payload: SyncDocumentStatePayload) => void;
    onMenuAction: (callback: (action: MenuAction) => void | Promise<void>) => () => void;
  };
}
