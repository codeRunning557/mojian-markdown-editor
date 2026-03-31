/// <reference types="vite/client" />

type MenuAction =
  | 'new-file'
  | 'open-file'
  | 'save-file'
  | 'save-file-as'
  | 'reference-image'
  | 'ai-beautify'
  | 'export-html'
  | 'export-pdf'
  | 'theme-paper'
  | 'theme-mist'
  | 'theme-slate'
  | 'theme-graphite'
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

type ReadFileResult = {
  filePath: string;
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

interface Window {
  markdownApp?: {
    ping: () => string;
    openMarkdownFile: () => Promise<OpenFileResult>;
    saveMarkdownFile: (payload: SaveFilePayload) => Promise<SaveFileResult>;
    readMarkdownFile: (filePath: string) => Promise<ReadFileResult>;
    exportHtmlFile: (payload: ExportHtmlPayload) => Promise<SaveFileResult>;
    exportPdfFile: (payload: ExportPdfPayload) => Promise<SaveFileResult>;
    saveImageAsset: (payload: SaveImageAssetPayload) => Promise<SaveFileResult>;
    pickImageReference: (payload: PickImageReferencePayload) => Promise<PickImageReferenceResult>;
    onMenuAction: (callback: (action: MenuAction) => void | Promise<void>) => () => void;
  };
}
