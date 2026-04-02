
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type ClipboardEvent,
  type DragEvent,
  type FormEvent,
  type KeyboardEvent,
  type MouseEvent as ReactMouseEvent
} from 'react';
import { createPortal } from 'react-dom';
import DOMPurify from 'dompurify';
import { marked } from 'marked';

type ViewMode = 'write' | 'split' | 'preview';
type SelectionSource = 'editor' | 'preview';
type BeautifyStrength = 'light' | 'standard' | 'deep';
type AIProvider = 'builtin' | 'openai-compatible';
type AIContextScope = 'selection' | 'section' | 'document';
type AIActionKind = 'rewrite' | 'expand' | 'shorten' | 'continue';
type AIPresetKey = 'builtin' | 'siliconflow' | 'dashscope' | 'deepseek' | 'hunyuan' | 'custom';

type EditorFileState = {
  filePath: string | null;
  displayName: string;
  content: string;
};

type RecentFile = {
  filePath: string;
  displayName: string;
};

type OutlineItem = {
  id: string;
  level: number;
  text: string;
  start: number;
};

type HeadingInfo = {
  level: number;
  text: string;
};

type BeautifyResult = {
  content: string;
  summary: string;
  changes: string[];
};

type DiffRow = {
  left: string;
  right: string;
  type: 'same' | 'changed' | 'added' | 'removed';
};

type BeautifyPreviewState = {
  scope: 'document' | 'selection';
  strength: BeautifyStrength;
  original: string;
  beautified: string;
  summary: string;
  changes: string[];
  hasChanges: boolean;
  selectionRange: { start: number; end: number } | null;
};

type ContextMenuState = {
  x: number;
  y: number;
};

type ContextMenuAction = {
  key: string;
  label: string;
  onClick: () => void | Promise<void>;
  active?: boolean;
};

type ContextMenuGroup = {
  key: string;
  title: string;
  columns?: 2 | 3;
  items: ContextMenuAction[];
};

type TranslationDirection = 'zh-to-en' | 'en-to-zh';

type AISettings = {
  provider: AIProvider;
  apiBase: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxTokens: number;
};

type AIProfile = AISettings & {
  id: string;
  name: string;
};

type TranslationPreviewState = {
  scope: 'document' | 'selection';
  direction: TranslationDirection;
  original: string;
  translated: string;
  providerLabel: string;
  hasChanges: boolean;
  selectionRange: { start: number; end: number } | null;
};

type DocxImportPreviewState = {
  sourceFilePath: string;
  sourceFileName: string;
  markdown: string;
  previewHtml: string;
  messages: string[];
  summary: DocxImportSummary;
  imageAssets: DocxImportImageAsset[];
};

type AIChatMessage = {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  contextScope: AIContextScope;
  providerLabel?: string;
  selectionRange?: { start: number; end: number } | null;
  sectionId?: string | null;
  sectionRange?: { start: number; end: number } | null;
};

type AIActionPreviewState = {
  action: AIActionKind;
  scope: AIContextScope;
  original: string;
  result: string;
  providerLabel: string;
  hasChanges: boolean;
  sectionId: string | null;
  selectionRange: { start: number; end: number } | null;
  sectionRange: { start: number; end: number } | null;
};

type AIRequestLog = {
  id: string;
  time: string;
  kind: 'chat' | 'action' | 'translate' | 'connection';
  status: 'started' | 'success' | 'error' | 'skipped';
  providerLabel: string;
  endpoint: string;
  detail: string;
};

type ThemeName = 'paper' | 'mist' | 'slate' | 'graphite' | 'terminal' | 'nightcode' | 'campus' | 'youth';
type EditorHistoryEntry = {
  content: string;
  selection: { start: number; end: number } | null;
};
type DocumentHistoryEntry = {
  content: string;
  selection: { start: number; end: number } | null;
};

const recentFilesStorageKey = 'paperflow.recent-files';
const aiSettingsStorageKey = 'paperflow.ai-settings';
const aiProfilesStorageKey = 'paperflow.ai-profiles';
const aiActiveProfileStorageKey = 'paperflow.ai-active-profile';
const legacyTranslationSettingsStorageKey = 'paperflow.translation-settings';
const themeStorageKey = 'paperflow.theme';
const allowedUriPattern = /^(?:(?:https?|file|mailto|tel|data):|[^a-z]|[a-z+.\-]+(?:[^a-z+.\-:]|$))/i;
const aiRequestTimeoutMs = 180000;
const aiProviderPresets = {
  builtin: {
    label: '内置文档助手',
    provider: 'builtin',
    apiBase: '',
    model: '',
    description: '不联网，可直接使用，能力较基础。',
    consoleUrl: ''
  },
  siliconflow: {
    label: 'SiliconFlow',
    provider: 'openai-compatible',
    apiBase: 'https://api.siliconflow.cn/v1',
    model: 'Qwen/Qwen2-7B-Instruct',
    description: '国内优先推荐，适合先跑通测试。',
    consoleUrl: 'https://cloud.siliconflow.cn/'
  },
  dashscope: {
    label: '阿里云百炼',
    provider: 'openai-compatible',
    apiBase: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    model: 'qwen-plus',
    description: '国内大厂接口，适合稳定使用。',
    consoleUrl: 'https://bailian.console.aliyun.com/'
  },
  deepseek: {
    label: 'DeepSeek',
    provider: 'openai-compatible',
    apiBase: 'https://api.deepseek.com/v1',
    model: 'deepseek-chat',
    description: '对话和改写能力强，性价比较高。',
    consoleUrl: 'https://platform.deepseek.com/'
  },
  hunyuan: {
    label: '腾讯混元',
    provider: 'openai-compatible',
    apiBase: 'https://api.hunyuan.cloud.tencent.com/v1',
    model: 'hunyuan-turbos-latest',
    description: '国内可选预设，适合补充测试。',
    consoleUrl: 'https://console.cloud.tencent.com/hunyuan'
  },
  custom: {
    label: '自定义兼容接口',
    provider: 'openai-compatible',
    apiBase: '',
    model: '',
    description: '手动填写 OpenAI-compatible 接口地址和模型名。',
    consoleUrl: ''
  }
} as const satisfies Record<
  AIPresetKey,
  { label: string; provider: AIProvider; apiBase: string; model: string; description: string; consoleUrl: string }
>;

const starterDocument = `# 欢迎使用墨笺

这是一款面向本地写作的 Markdown 编辑器。

- 点击顶部按钮或菜单打开 \`.md\` 文档
- 左侧编辑，右侧实时预览
- 支持图片插入、代码复制和 AI 美化

## 快速开始

1. 点击“新建”
2. 输入一段 Markdown
3. 点击“AI 美化”预览优化结果

## 代码块示例

\`\`\`bash
npm run dev
\`\`\`
`;

marked.setOptions({
  gfm: true,
  breaks: true
});

function normalizeLineEndings(value: string) {
  return value.replace(/\r\n/g, '\n');
}

function sanitizeHtml(html: string) {
  return DOMPurify.sanitize(html, {
    ALLOWED_URI_REGEXP: allowedUriPattern
  });
}

function escapeHtml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function normalizeExplicitLanguage(language: string) {
  const normalized = language.trim().toLowerCase();
  if (!normalized) {
    return '';
  }

  const aliases = new Map([
    ['ps', 'powershell'],
    ['pwsh', 'powershell'],
    ['shell', 'bash'],
    ['sh', 'bash'],
    ['zsh', 'bash'],
    ['cmd', 'batch'],
    ['bat', 'batch'],
    ['js', 'javascript'],
    ['ts', 'typescript'],
    ['yml', 'yaml'],
    ['py', 'python'],
    ['cs', 'csharp']
  ]);

  return aliases.get(normalized) ?? normalized;
}

function renderMarkdown(content: string, withCodeToolbar: boolean) {
  const renderer = new marked.Renderer();

  renderer.code = ({ text, lang }) => {
    const language = normalizeExplicitLanguage((lang || '').trim()) || 'text';
    const languageLabel = language === 'text' ? '纯文本' : language;
    const escapedCode = escapeHtml(text);
    const escapedLanguage = escapeHtml(language);
    const escapedLanguageLabel = escapeHtml(languageLabel);

    if (!withCodeToolbar) {
      return `<pre><code class="language-${escapedLanguage}">${escapedCode}</code></pre>`;
    }

    return `
      <div class="code-block">
        <div class="code-block-toolbar">
          <span class="code-block-language">${escapedLanguageLabel}</span>
          <button type="button" class="code-copy-button">复制</button>
        </div>
        <pre><code class="language-${escapedLanguage}">${escapedCode}</code></pre>
      </div>
    `;
  };

  return marked.parse(content, { renderer }) as string;
}

function extractOutlineFromMarkdown(content: string): OutlineItem[] {
  const lines = normalizeLineEndings(content).split('\n');
  const outline: OutlineItem[] = [];
  let headingIndex = 0;
  let inFence = false;
  let fenceMarker = '';
  let offset = 0;

  for (const line of lines) {
    const fenceMatch = line.match(/^(\s*)(`{3,}|~{3,})/);
    if (fenceMatch) {
      const marker = fenceMatch[2][0];
      if (!inFence) {
        inFence = true;
        fenceMarker = marker;
      } else if (fenceMarker === marker) {
        inFence = false;
        fenceMarker = '';
      }
      offset += line.length + 1;
      continue;
    }

    if (inFence) {
      offset += line.length + 1;
      continue;
    }

    const headingMatch = line.match(/^(#{1,6})\s+(.*)$/);
    if (!headingMatch) {
      offset += line.length + 1;
      continue;
    }

    outline.push({
      id: `heading-${headingIndex}`,
      level: headingMatch[1].length,
      text: headingMatch[2].trim(),
      start: offset
    });
    headingIndex += 1;
    offset += line.length + 1;
  }

  return outline;
}

function buildEditableHtml(content: string, outline: OutlineItem[]) {
  const rawHtml = sanitizeHtml(renderMarkdown(content, false));
  const parser = new DOMParser();
  const documentNode = parser.parseFromString(`<div data-editor-root="true">${rawHtml}</div>`, 'text/html');
  const root = documentNode.body.firstElementChild as HTMLElement | null;
  if (!root) {
    return rawHtml;
  }

  let headingIndex = 0;
  let currentSectionId: string | null = null;

  Array.from(root.children).forEach((element) => {
    const headingTag = /^H[1-6]$/.test(element.tagName);

    if (headingTag) {
      const id = outline[headingIndex]?.id ?? `heading-${headingIndex}`;
      headingIndex += 1;
      element.setAttribute('data-outline-id', id);
      element.setAttribute('data-section-id', id);
      currentSectionId = id;
      return;
    }

    if (currentSectionId) {
      element.setAttribute('data-section-id', currentSectionId);
    }
  });

  return root.innerHTML;
}

function serializeInlineMarkdown(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return node.textContent?.replace(/\u00a0/g, ' ') ?? '';
  }

  if (!(node instanceof HTMLElement)) {
    return '';
  }

  const content = Array.from(node.childNodes).map(serializeInlineMarkdown).join('');
  const tagName = node.tagName.toLowerCase();

  if (tagName === 'br') {
    return '\n';
  }

  if (tagName === 'strong' || tagName === 'b') {
    return `**${content}**`;
  }

  if (tagName === 'em' || tagName === 'i') {
    return `*${content}*`;
  }

  if (tagName === 'code' && node.parentElement?.tagName.toLowerCase() !== 'pre') {
    return `\`${content}\``;
  }

  if (tagName === 'a') {
    const href = node.getAttribute('href') ?? '';
    return href ? `[${content}](${href})` : content;
  }

  if (tagName === 'img') {
    const alt = node.getAttribute('alt') ?? 'image';
    const src = node.getAttribute('src') ?? '';
    return src ? `![${alt}](${src})` : '';
  }

  return content;
}

function serializePreMarkdown(element: HTMLElement) {
  const code = element.querySelector('code');
  const className = code?.className ?? '';
  const languageMatch = className.match(/language-([\w-]+)/);
  const language = normalizeExplicitLanguage(languageMatch?.[1] ?? '');
  const codeText = code?.textContent ?? element.textContent ?? '';
  const fence = language ? `\`\`\`${language}` : '```';
  return `${fence}\n${codeText.replace(/\n+$/, '')}\n\`\`\``;
}

function serializeListMarkdown(element: HTMLElement, ordered: boolean, depth = 0): string {
  const items = Array.from(element.children).filter((child): child is HTMLElement => child.tagName.toLowerCase() === 'li');

  return items
    .map((item, index) => {
      const prefix = ordered ? `${index + 1}. ` : '- ';
      const inlineParts: string[] = [];
      const nestedBlocks: string[] = [];

      Array.from(item.childNodes).forEach((child) => {
        if (child instanceof HTMLElement && (child.tagName === 'UL' || child.tagName === 'OL')) {
          nestedBlocks.push(serializeListMarkdown(child, child.tagName === 'OL', depth + 1));
          return;
        }

        if (child instanceof HTMLElement && child.tagName === 'PRE') {
          nestedBlocks.push(serializePreMarkdown(child));
          return;
        }

        inlineParts.push(serializeInlineMarkdown(child));
      });

      const firstLine = `${'  '.repeat(depth)}${prefix}${inlineParts.join('').trim()}`.trimEnd();
      const extra = nestedBlocks
        .filter(Boolean)
        .map((block) =>
          block
            .split('\n')
            .map((line) => `${'  '.repeat(depth + 1)}${line}`)
            .join('\n')
        );

      return [firstLine, ...extra].filter(Boolean).join('\n');
    })
    .join('\n');
}

function serializeBlockMarkdown(node: Node): string[] {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent?.replace(/\u00a0/g, ' ').trim() ?? '';
    return text ? [text] : [];
  }

  if (!(node instanceof HTMLElement)) {
    return [];
  }

  const tagName = node.tagName.toLowerCase();

  if (/^h[1-6]$/.test(tagName)) {
    const level = Number.parseInt(tagName[1] ?? '1', 10);
    const text = Array.from(node.childNodes).map(serializeInlineMarkdown).join('').trim();
    return text ? [`${'#'.repeat(level)} ${text}`] : [];
  }

  if (tagName === 'p') {
    const text = Array.from(node.childNodes).map(serializeInlineMarkdown).join('').trim();
    return text ? [text] : [];
  }

  if (tagName === 'pre') {
    return [serializePreMarkdown(node)];
  }

  if (tagName === 'ul' || tagName === 'ol') {
    return [serializeListMarkdown(node, tagName === 'ol')];
  }

  if (tagName === 'blockquote') {
    const blocks = Array.from(node.childNodes).flatMap(serializeBlockMarkdown);
    const content = blocks.join('\n\n');
    return content
      ? [
          content
            .split('\n')
            .map((line) => (line ? `> ${line}` : '>'))
            .join('\n')
        ]
      : [];
  }

  if (tagName === 'hr') {
    return ['---'];
  }

  if (tagName === 'img') {
    return [serializeInlineMarkdown(node)];
  }

  if (tagName === 'div') {
    const childBlocks = Array.from(node.childNodes).flatMap(serializeBlockMarkdown);
    if (childBlocks.length > 0) {
      return childBlocks;
    }

    const text = Array.from(node.childNodes).map(serializeInlineMarkdown).join('').trim();
    return text ? [text] : [];
  }

  const fallbackText = Array.from(node.childNodes).map(serializeInlineMarkdown).join('').trim();
  return fallbackText ? [fallbackText] : [];
}

function serializeRichEditorContent(root: HTMLElement) {
  return Array.from(root.childNodes)
    .flatMap(serializeBlockMarkdown)
    .map((block) => block.trimEnd())
    .filter((block, index, all) => block.length > 0 || (index > 0 && all[index - 1].length > 0))
    .join('\n\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function setContentEditableSelection(root: HTMLElement, start: number, end = start) {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let current = 0;
  let startNode: Node | null = null;
  let endNode: Node | null = null;
  let startOffset = 0;
  let endOffset = 0;

  while (walker.nextNode()) {
    const node = walker.currentNode;
    const length = node.textContent?.length ?? 0;

    if (!startNode && start <= current + length) {
      startNode = node;
      startOffset = Math.max(0, start - current);
    }

    if (!endNode && end <= current + length) {
      endNode = node;
      endOffset = Math.max(0, end - current);
      break;
    }

    current += length;
  }

  const range = document.createRange();
  range.setStart(startNode ?? root, startNode ? startOffset : root.childNodes.length);
  range.setEnd(endNode ?? startNode ?? root, endNode ? endOffset : startNode ? startOffset : root.childNodes.length);
  selection.removeAllRanges();
  selection.addRange(range);
}

function getContentEditableRangeOffsets(root: HTMLElement, range: Range) {
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) {
    return null;
  }

  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let current = 0;
  let start: number | null = null;
  let end: number | null = null;

  while (walker.nextNode()) {
    const node = walker.currentNode;
    const length = node.textContent?.length ?? 0;

    if (node === range.startContainer) {
      start = current + range.startOffset;
    }

    if (node === range.endContainer) {
      end = current + range.endOffset;
      break;
    }

    current += length;
  }

  if (start === null || end === null) {
    return null;
  }

  return { start, end };
}

function getContentEditableSelectionOffsets(root: HTMLElement) {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return null;
  }

  const range = selection.getRangeAt(0);
  return getContentEditableRangeOffsets(root, range);
}

function applyDomRangeSelection(range: Range) {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  selection.removeAllRanges();
  selection.addRange(range);
}

function resolveRangeFromPoint(x: number, y: number) {
  const positionApi = document.caretPositionFromPoint as
    | ((x: number, y: number) => { offsetNode: Node; offset: number } | null)
    | undefined;
  if (positionApi) {
    const position = positionApi(x, y);
    if (!position) {
      return null;
    }
    const range = document.createRange();
    range.setStart(position.offsetNode, position.offset);
    range.collapse(true);
    return range;
  }

  const rangeApi = document.caretRangeFromPoint as ((x: number, y: number) => Range | null) | undefined;
  return rangeApi?.(x, y) ?? null;
}

function moveCaretToNearestLineEnd(root: HTMLElement, target: EventTarget | null, x: number, y: number) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }

  const block = target.closest('li, p, h1, h2, h3, h4, h5, h6, blockquote, pre, div');
  if (!block || !root.contains(block)) {
    return false;
  }

  const blockRange = document.createRange();
  blockRange.selectNodeContents(block);
  const rects = Array.from(blockRange.getClientRects());
  const lineRect = rects.find((rect) => y >= rect.top - 2 && y <= rect.bottom + 2);

  if (!lineRect || x <= lineRect.right + 2) {
    return false;
  }

  const caretRange = resolveRangeFromPoint(Math.max(lineRect.left + 1, lineRect.right - 2), y);
  if (!caretRange || !root.contains(caretRange.startContainer)) {
    return false;
  }

  applyDomRangeSelection(caretRange);
  return true;
}

function getClosestEditorBlock(root: HTMLElement, node: Node | null) {
  if (!node) {
    return null;
  }

  const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  const block = element?.closest<HTMLElement>(
    '[data-outline-id], p, li, pre, blockquote, table, hr, ul, ol, div, h1, h2, h3, h4, h5, h6'
  );

  if (!block || !root.contains(block)) {
    return null;
  }

  return block;
}

function getRootLevelEditorBlock(root: HTMLElement, node: Node | null) {
  let block = getClosestEditorBlock(root, node);
  while (block && block.parentElement && block.parentElement !== root) {
    block = block.parentElement;
  }
  return block && block.parentElement === root ? block : null;
}

function formatFileName(filePath: string | null, displayName?: string) {
  if (displayName) {
    return displayName;
  }

  if (!filePath) {
    return '未命名.md';
  }

  const segments = filePath.split(/[\\/]/);
  return segments[segments.length - 1] ?? '未命名.md';
}

function loadRecentFiles(): RecentFile[] {
  try {
    const raw = window.localStorage.getItem(recentFilesStorageKey);
    if (!raw) {
      return [];
    }

    const parsed = JSON.parse(raw) as RecentFile[];
    return parsed.filter((item) => item.filePath && item.displayName).slice(0, 8);
  } catch {
    return [];
  }
}

function normalizeAiSettings(parsed: Partial<AISettings> | null | undefined): AISettings {
  return {
    provider: parsed?.provider === 'openai-compatible' ? 'openai-compatible' : 'builtin',
    apiBase: parsed?.apiBase ?? '',
    apiKey: parsed?.apiKey ?? '',
    model: parsed?.model ?? '',
    temperature:
      typeof parsed?.temperature === 'number' && Number.isFinite(parsed.temperature)
        ? Math.min(1.2, Math.max(0, parsed.temperature))
        : 0.2,
    maxTokens:
      typeof parsed?.maxTokens === 'number' && Number.isFinite(parsed.maxTokens)
        ? Math.min(4000, Math.max(200, Math.round(parsed.maxTokens)))
        : 1200
  };
}

function createAiProfileId() {
  return `profile-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function createAiProfile(
  name: string,
  settings?: Partial<AISettings>,
  options?: { id?: string }
): AIProfile {
  return {
    id: options?.id ?? createAiProfileId(),
    name,
    ...normalizeAiSettings(settings)
  };
}

function normalizeAiProfileName(name: string | null | undefined) {
  return String(name || '').trim() || '未命名模型';
}

function normalizeAiProfilesForPersistence(profiles: AIProfile[]): AIProfile[] {
  const profileMap = new Map<string, AIProfile>();

  profiles.forEach((profile) => {
    const normalizedName = normalizeAiProfileName(profile.name);
    if (profileMap.has(normalizedName)) {
      profileMap.delete(normalizedName);
    }

    profileMap.set(normalizedName, {
      ...profile,
      name: normalizedName
    });
  });

  return Array.from(profileMap.values());
}

function toPersistedAiProfiles(profiles: AIProfile[]): PersistedAIProfile[] {
  return normalizeAiProfilesForPersistence(profiles).map((profile) => ({
    name: profile.name,
    provider: profile.provider,
    apiBase: profile.apiBase,
    apiKey: profile.apiKey,
    model: profile.model,
    temperature: profile.temperature,
    maxTokens: profile.maxTokens
  }));
}

function buildAiProfilesFromPersistedConfig(profiles: PersistedAIProfile[]): AIProfile[] {
  const normalizedProfiles = normalizeAiProfilesForPersistence(
    profiles.map((profile) => createAiProfile(normalizeAiProfileName(profile.name), profile))
  );

  return normalizedProfiles.length > 0 ? normalizedProfiles : [createAiProfile('内置文档助手')];
}

function loadLegacyAiSettings(): AISettings {
  try {
    const raw =
      window.localStorage.getItem(aiSettingsStorageKey) ??
      window.localStorage.getItem(legacyTranslationSettingsStorageKey);
    if (!raw) {
      return normalizeAiSettings(null);
    }

    const parsed = JSON.parse(raw) as Partial<AISettings>;
    return normalizeAiSettings({ ...parsed, apiKey: '' });
  } catch {
    return normalizeAiSettings(null);
  }
}

function loadAiProfiles(): AIProfile[] {
  try {
    const raw = window.localStorage.getItem(aiProfilesStorageKey);
    if (raw) {
      const parsed = JSON.parse(raw) as Array<Partial<AIProfile>>;
      const profiles = normalizeAiProfilesForPersistence(
        parsed
        .filter((item) => typeof item?.id === 'string' && typeof item?.name === 'string')
        .map((item) => ({
          id: item.id as string,
          name: normalizeAiProfileName(item.name as string),
          ...normalizeAiSettings({ ...item, apiKey: '' })
        }))
      );

      if (profiles.length > 0) {
        return profiles;
      }
    }
  } catch {
    // ignore and fallback to migration
  }

  const legacy = loadLegacyAiSettings();
  const legacyPreset = inferAiPreset(legacy);
  const legacyName =
    legacy.provider === 'builtin'
      ? '内置文档助手'
      : legacyPreset === 'custom'
        ? '自定义兼容模型'
        : aiProviderPresets[legacyPreset].label;

  return [createAiProfile(legacyName, legacy)];
}

function toLocalSafeAiProfiles(profiles: AIProfile[]) {
  return normalizeAiProfilesForPersistence(profiles).map((profile) => ({
    id: profile.id,
    name: profile.name,
    provider: profile.provider,
    apiBase: profile.apiBase,
    model: profile.model,
    temperature: profile.temperature,
    maxTokens: profile.maxTokens
  }));
}

function loadActiveAiProfileId(profiles: AIProfile[]) {
  try {
    const raw = window.localStorage.getItem(aiActiveProfileStorageKey);
    if (raw && profiles.some((profile) => profile.id === raw)) {
      return raw;
    }
  } catch {
    // ignore
  }

  return profiles[0]?.id ?? '';
}

function loadThemePreference(): ThemeName {
  try {
    const raw = window.localStorage.getItem(themeStorageKey);
    if (
      raw === 'mist' ||
      raw === 'slate' ||
      raw === 'graphite' ||
      raw === 'terminal' ||
      raw === 'nightcode' ||
      raw === 'campus' ||
      raw === 'youth'
    ) {
      return raw;
    }
  } catch {
    return 'paper';
  }

  return 'paper';
}

function inferAiPreset(settings: AISettings): AIPresetKey {
  if (settings.provider === 'builtin') {
    return 'builtin';
  }

  const apiBase = settings.apiBase.trim().replace(/\/$/, '');
  const model = settings.model.trim();

  if (
    apiBase === aiProviderPresets.siliconflow.apiBase &&
    model === aiProviderPresets.siliconflow.model
  ) {
    return 'siliconflow';
  }

  if (
    apiBase === aiProviderPresets.dashscope.apiBase &&
    model === aiProviderPresets.dashscope.model
  ) {
    return 'dashscope';
  }

  if (apiBase === aiProviderPresets.deepseek.apiBase && model === aiProviderPresets.deepseek.model) {
    return 'deepseek';
  }

  if (apiBase === aiProviderPresets.hunyuan.apiBase && model === aiProviderPresets.hunyuan.model) {
    return 'hunyuan';
  }

  return 'custom';
}

function buildExportHtml(title: string, bodyHtml: string) {
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${title}</title>
    <style>
      body {
        max-width: 840px;
        margin: 0 auto;
        padding: 40px 24px 80px;
        font-family: "Segoe UI", "Microsoft YaHei", sans-serif;
        line-height: 1.8;
        color: #2f2418;
        background: #fbf6ee;
      }
      h1, h2, h3, h4, h5, h6 { line-height: 1.25; color: #3a2712; }
      h1 { font-size: 2.25rem; font-weight: 800; }
      h2 { font-size: 1.7rem; font-weight: 760; }
      h3 { font-size: 1.35rem; font-weight: 700; }
      pre { overflow: auto; padding: 16px; border-radius: 12px; background: #2f2418; color: #f6e6d1; }
      code { padding: 0.1em 0.35em; border-radius: 6px; background: rgba(142, 90, 36, 0.1); }
      pre code { padding: 0; background: transparent; }
      blockquote { margin: 0; padding-left: 16px; border-left: 4px solid rgba(142, 90, 36, 0.3); color: #5e4d3c; }
      table { width: 100%; border-collapse: collapse; }
      th, td { padding: 10px 12px; border: 1px solid rgba(103, 71, 33, 0.14); }
      img { max-width: 100%; }
    </style>
  </head>
  <body>
    ${bodyHtml}
  </body>
</html>`;
}

function getImageMarkdown(relativePath: string, altText = 'image') {
  const normalized = relativePath.replace(/\\/g, '/');
  return `![${altText}](${normalized})`;
}

function getDisplayNameWithoutExtension(fileName: string) {
  return fileName.replace(/\.[^.]+$/, '') || 'image';
}

function inferCodeLanguage(code: string) {
  const text = code.trim();
  const lower = text.toLowerCase();

  if (!text) {
    return 'text';
  }

  if (
    /^docker compose\b/m.test(lower) ||
    /^docker\b/m.test(lower) ||
    /^sudo\b/m.test(lower) ||
    /^chmod\b/m.test(lower) ||
    /^npm\b/m.test(lower) ||
    /^pnpm\b/m.test(lower) ||
    /^yarn\b/m.test(lower) ||
    /^git\b/m.test(lower) ||
    /^curl\b/m.test(lower)
  ) {
    return 'bash';
  }

  if (
    /^\s*get-[a-z]/im.test(text) ||
    /^\s*set-[a-z]/im.test(text) ||
    /^\s*new-[a-z]/im.test(text) ||
    /\$env:/i.test(text) ||
    /write-host/i.test(text) ||
    /set-location/i.test(text)
  ) {
    return 'powershell';
  }

  if (/^\s*[{[]/.test(text) && /"\s*:\s*/.test(text)) {
    return 'json';
  }

  if (/^\s*[a-z0-9_.-]+\s*=\s*.+/im.test(text) && /\[[^\]]+\]/.test(text)) {
    return 'toml';
  }

  if (/^\s*[a-z0-9_.-]+\s*=\s*.+/im.test(text) && !/[{}\[\]]/.test(text)) {
    return 'ini';
  }

  if (/^\s*[a-z0-9_-]+\s*:\s+/im.test(text) && !/[;{}()]/.test(text)) {
    return 'yaml';
  }

  if (/<[a-z][\s\S]*>/i.test(text)) {
    return /<(html|body|div|span|script|style)\b/i.test(text) ? 'html' : 'xml';
  }

  if (/^\s*FROM\s+\S+/im.test(text) || /^\s*RUN\s+/im.test(text) || /^\s*CMD\s+/im.test(text)) {
    return 'dockerfile';
  }

  if (/\bselect\b[\s\S]+\bfrom\b/i.test(text) || /\binsert into\b/i.test(text) || /\bcreate table\b/i.test(text)) {
    return 'sql';
  }

  if (/\bpublic\s+class\b/.test(text) || /\bSystem\.out\.println\b/.test(text) || /\b@SpringBootApplication\b/.test(text)) {
    return 'java';
  }

  if (/\busing\s+System\b/.test(text) || /\bnamespace\s+\w+/.test(text) || /\bConsole\.WriteLine\b/.test(text)) {
    return 'csharp';
  }

  if (/\bdef\s+\w+\(/.test(text) || /\bimport\s+\w+/.test(text) || /\bprint\(/.test(text)) {
    return 'python';
  }

  if (/\binterface\s+\w+/.test(text) || /:\s*(string|number|boolean|unknown|any)([,\)\];]|$)/.test(text) || /\btype\s+\w+\s*=/.test(text)) {
    return 'typescript';
  }

  if (/\bconst\b|\blet\b|\bfunction\b|=>/.test(text)) {
    return 'javascript';
  }

  if (/\bpackage\s+main\b/.test(text) || /\bfunc\s+\w+\(/.test(text)) {
    return 'go';
  }

  if (/\bfn\s+main\s*\(/.test(text) || /\blet\s+mut\b/.test(text)) {
    return 'rust';
  }

  if (/^\s*#/m.test(text) && /\binclude\b/.test(text)) {
    return 'cpp';
  }

  if (/^\s*\.[\w-]+\s*{/m.test(text) || /@media\s*\(/.test(text)) {
    return 'css';
  }

  if (/^\s*server\s*{/m.test(text) || /\bproxy_pass\b/.test(text)) {
    return 'nginx';
  }

  return 'text';
}

function extractHeadings(lines: string[]) {
  return lines
    .map((line) => line.match(/^(#{1,6})\s+(.+)$/))
    .filter((match): match is RegExpMatchArray => Boolean(match))
    .map((match) => ({
      level: match[1].length,
      text: match[2].trim()
    }))
    .filter((item) => !/^目录$/i.test(item.text));
}

function shouldIncludeHeadingInToc(heading: HeadingInfo, tocStrength: number) {
  const text = heading.text.trim();

  if (isMinorStepHeading(text)) {
    return false;
  }

  if (tocStrength <= 1) {
    return heading.level === 1;
  }

  if (tocStrength === 2) {
    return heading.level === 1 || (heading.level === 2 && isFrameworkHeading(text));
  }

  if (tocStrength === 3) {
    return heading.level <= 2;
  }

  if (tocStrength === 4) {
    return heading.level <= 2 || (heading.level === 3 && isFrameworkHeading(text));
  }

  return heading.level <= 3;
}

function buildTableOfContents(headings: HeadingInfo[], tocStrength: number) {
  const frameworkHeadings = headings.filter((heading) => shouldIncludeHeadingInToc(heading, tocStrength));

  if (frameworkHeadings.length < 2) {
    return null;
  }

  const tocLines = ['## 目录', ''];
  for (const heading of frameworkHeadings) {
    const indent = '  '.repeat(Math.max(0, heading.level - 1));
    tocLines.push(`${indent}- ${heading.text}`);
  }
  tocLines.push('');
  return tocLines;
}

function analyzeWritingHabits(content: string) {
  const bulletMarks = Array.from(content.matchAll(/^\s*([-*+])\s+/gm)).map((match) => match[1]);

  return {
    bulletPreference: bulletMarks[0] ?? '-'
  };
}

function buildDiffRows(original: string, beautified: string) {
  const leftLines = normalizeLineEndings(original).split('\n');
  const rightLines = normalizeLineEndings(beautified).split('\n');
  const rows: DiffRow[] = [];
  const total = Math.max(leftLines.length, rightLines.length);

  for (let index = 0; index < total; index += 1) {
    const left = leftLines[index] ?? '';
    const right = rightLines[index] ?? '';
    let type: DiffRow['type'] = 'same';

    if (index >= leftLines.length) {
      type = 'added';
    } else if (index >= rightLines.length) {
      type = 'removed';
    } else if (left !== right) {
      type = 'changed';
    }

    rows.push({ left, right, type });
  }

  return rows;
}

function isBlankLine(line: string) {
  return line.trim() === '';
}

function isSpecialBlockLine(line: string) {
  const text = line.trim();

  return (
    text === '' ||
    /^```/.test(text) ||
    /^[-*+]\s+/.test(text) ||
    /^\d+[.)]\s+/.test(text) ||
    /^\|/.test(text) ||
    /^>/.test(text) ||
    /^!\[/.test(text) ||
    /^<.+>$/.test(text) ||
    /^-{3,}$/.test(text)
  );
}

function isLikelySentence(line: string) {
  const text = line.trim();
  return text.length > 34 || /[。！？；：:，,、；.]$/.test(text);
}

function isMinorStepHeading(text: string) {
  const normalized = text.trim();

  return (
    /^\d+[.)]\s+/.test(normalized) ||
    /^\d+\.\d+(\.\d+)?\s+/.test(normalized) ||
    /^[（(][一二三四五六七八九十0-9]+[)）]\s*/.test(normalized) ||
    /^\d+[)）]\s+/.test(normalized) ||
    /^[①②③④⑤⑥⑦⑧⑨⑩]/.test(normalized) ||
    /^第[一二三四五六七八九十百千万0-9]+步/.test(normalized)
  );
}

function isFrameworkHeading(text: string) {
  const normalized = text.trim();

  return (
    /^#\s+/.test(normalized) ||
    /^第[一二三四五六七八九十百千万0-9]+(章|节|部分|篇)/.test(normalized) ||
    /^[一二三四五六七八九十]+[、.．]\s*/.test(normalized) ||
    /^(前言|引言|摘要|概述|背景|结论|总结|附录|常见问题|FAQ|部署|使用方法|注意事项|问题背景|问题分析|解决方案|安装步骤|快速开始)\b/.test(normalized)
  );
}

function demoteMinorStepHeading(text: string) {
  const normalized = text.trim();

  if (/^\d+[.)]\s+/.test(normalized)) {
    return normalized.replace(/^(\d+)[)]\s+/, '$1. ');
  }

  return normalized;
}

function buildBuiltinGlossary(direction: TranslationDirection) {
  if (direction === 'zh-to-en') {
    return [
      ['快速开始', 'Quick Start'],
      ['解决方案', 'Solution'],
      ['问题背景', 'Background'],
      ['问题分析', 'Analysis'],
      ['注意事项', 'Notes'],
      ['部署', 'Deployment'],
      ['安装', 'Installation'],
      ['步骤', 'Steps'],
      ['示例', 'Example'],
      ['目录', 'Table of Contents'],
      ['总结', 'Summary'],
      ['结论', 'Conclusion'],
      ['用户', 'user'],
      ['用户名', 'username'],
      ['命令', 'command'],
      ['输入', 'enter'],
      ['打开', 'open'],
      ['保存', 'save'],
      ['文件', 'file'],
      ['图片', 'image'],
      ['复制', 'copy'],
      ['预览', 'preview'],
      ['终端', 'terminal'],
      ['管理员', 'administrator'],
      ['系统', 'system']
    ] as const;
  }

  return [
    ['Quick Start', '快速开始'],
    ['Solution', '解决方案'],
    ['Background', '问题背景'],
    ['Analysis', '问题分析'],
    ['Notes', '注意事项'],
    ['Deployment', '部署'],
    ['Installation', '安装'],
    ['Steps', '步骤'],
    ['Example', '示例'],
    ['Table of Contents', '目录'],
    ['Summary', '总结'],
    ['Conclusion', '结论'],
    ['username', '用户名'],
    ['user', '用户'],
    ['command', '命令'],
    ['enter', '输入'],
    ['open', '打开'],
    ['save', '保存'],
    ['file', '文件'],
    ['image', '图片'],
    ['copy', '复制'],
    ['preview', '预览'],
    ['terminal', '终端'],
    ['administrator', '管理员'],
    ['system', '系统']
  ] as const;
}

function translateBuiltinLine(text: string, direction: TranslationDirection) {
  let result = text;
  const glossary = buildBuiltinGlossary(direction);

  for (const [source, target] of glossary) {
    result = result.replaceAll(source, target);
  }

  if (direction === 'zh-to-en') {
    result = result
      .replace(/查看/g, 'Check')
      .replace(/使用/g, 'Use')
      .replace(/方法/g, 'method')
      .replace(/输入以下命令并回车[:：]?/g, 'Run the following command:')
      .replace(/按住/g, 'Press and hold ')
      .replace(/选择/g, 'select ')
      .replace(/即可/g, '')
      .replace(/。/g, '.');
  } else {
    result = result
      .replace(/\bPress and hold\b/g, '按住')
      .replace(/\bRun the following command:?/g, '输入以下命令并回车：')
      .replace(/\bselect\b/gi, '选择')
      .replace(/\bCheck\b/g, '查看')
      .replace(/\bUse\b/g, '使用');
  }

  return result;
}

function translateMarkdownWithBuiltin(content: string, direction: TranslationDirection) {
  const lines = normalizeLineEndings(content).split('\n');
  const output: string[] = [];
  let inCodeBlock = false;

  for (const line of lines) {
    if (/^```/.test(line.trim())) {
      inCodeBlock = !inCodeBlock;
      output.push(line);
      continue;
    }

    if (inCodeBlock || /^!\[/.test(line.trim())) {
      output.push(line);
      continue;
    }

    const match = line.match(/^(\s*(?:#{1,6}\s+|[-*+]\s+|\d+[.)]\s+|>\s+)?)?(.*)$/);
    const prefix = match?.[1] ?? '';
    const text = match?.[2] ?? line;
    output.push(`${prefix}${translateBuiltinLine(text, direction)}`);
  }

  return output.join('\n');
}

async function translateWithOpenAiCompatible(
  settings: AISettings,
  content: string,
  direction: TranslationDirection
) {
  const endpoint = settings.apiBase.replace(/\/$/, '');
  const response = await fetch(`${endpoint}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${settings.apiKey}`
    },
    body: JSON.stringify({
      model: settings.model,
      temperature: Math.min(settings.temperature, 0.2),
      max_tokens: settings.maxTokens,
      messages: [
        {
          role: 'system',
          content:
            direction === 'zh-to-en'
              ? 'Translate the markdown content from Simplified Chinese to English. Preserve markdown structure, headings, bullet markers, links, code blocks, inline code, and line breaks. Return only the translated markdown.'
              : 'Translate the markdown content from English to Simplified Chinese. Preserve markdown structure, headings, bullet markers, links, code blocks, inline code, and line breaks. Return only the translated markdown.'
        },
        {
          role: 'user',
          content
        }
      ]
    })
  });

  if (!response.ok) {
    throw new Error(`翻译接口返回 ${response.status}`);
  }

  const data = (await response.json()) as {
    choices?: Array<{ message?: { content?: string } }>;
  };
  const translated = data.choices?.[0]?.message?.content?.trim();
  if (!translated) {
    throw new Error('翻译接口没有返回内容');
  }

  return translated;
}

async function translateDocumentContent(
  content: string,
  direction: TranslationDirection,
  settings: AISettings
) {
  const canUseRemote =
    settings.provider === 'openai-compatible' &&
    settings.apiBase.trim() &&
    settings.apiKey.trim() &&
    settings.model.trim();

  if (canUseRemote) {
    const translated = await translateWithOpenAiCompatible(settings, content, direction);
    return {
      translated,
      providerLabel: '自定义兼容翻译 API'
    };
  }

  return {
    translated: translateMarkdownWithBuiltin(content, direction),
    providerLabel: '内置轻量翻译器'
  };
}

function buildBuiltinSummary(content: string) {
  const lines = normalizeLineEndings(content)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);
  const headings = extractHeadings(lines);
  const summaryLines: string[] = [];

  if (headings.length > 0) {
    summaryLines.push(`当前内容包含 ${headings.length} 个标题层级，主线如下：`);
    headings.slice(0, 6).forEach((heading, index) => {
      summaryLines.push(`${index + 1}. ${heading.text}`);
    });
  }

  const paragraphs = lines.filter((line) => !/^#{1,6}\s+/.test(line) && !/^[-*+]\s+/.test(line) && !/^\d+[.)]\s+/.test(line));
  if (paragraphs.length > 0) {
    summaryLines.push('');
    summaryLines.push('内容摘要：');
    summaryLines.push(paragraphs.slice(0, 3).map((line) => line.replace(/\s+/g, ' ')).join(' '));
  }

  return summaryLines.join('\n').trim() || '当前上下文内容较短，建议继续补充正文后再做总结。';
}

function normalizeTechnicalTone(text: string) {
  return text
    .replace(/你可以/g, '可')
    .replace(/然后/g, '随后')
    .replace(/这样就/g, '这样即可')
    .replace(/的话/g, '')
    .replace(/需要注意的是/g, '注意：')
    .replace(/搞定/g, '完成')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function buildBuiltinContinuation(text: string) {
  const trimmed = text.trim();
  const lines: string[] = [trimmed];

  if (/(部署|安装|启动|运行|配置|命令)/.test(trimmed)) {
    lines.push('');
    lines.push('建议继续补充以下信息：');
    lines.push('- 前置条件：说明执行前需要准备的环境或依赖');
    lines.push('- 验证方式：说明执行完成后如何确认结果正确');
    lines.push('- 常见问题：列出失败时的排查方向');
    return lines.join('\n');
  }

  lines.push('');
  lines.push('可以继续补充以下内容：');
  lines.push('- 背景与目标');
  lines.push('- 操作步骤或实现方式');
  lines.push('- 预期结果与注意事项');
  return lines.join('\n');
}

function transformWithBuiltinAi(action: AIActionKind, text: string) {
  const normalized = normalizeLineEndings(text).trim();
  if (!normalized) {
    return '';
  }

  if (action === 'rewrite') {
    return `${normalizeTechnicalTone(normalized)}\n`;
  }

  if (action === 'expand') {
    return `${normalizeTechnicalTone(normalized)}\n\n补充说明：\n- 建议明确目标与适用范围。\n- 建议补充关键步骤和验证方式。\n- 如涉及风险点，建议单独列出注意事项。\n`;
  }

  if (action === 'shorten') {
    const compact = normalized
      .replace(/为了能够/g, '为')
      .replace(/需要注意的是/g, '注意：')
      .replace(/在这种情况下/g, '此时')
      .replace(/\s+/g, ' ')
      .trim();
    return `${compact}\n`;
  }

  return `${buildBuiltinContinuation(normalized)}\n`;
}

function buildBuiltinAssistantReply(
  prompt: string,
  contextText: string,
  scope: AIContextScope,
  fileName: string
) {
  const normalizedPrompt = prompt.trim().toLowerCase();
  const summary = buildBuiltinSummary(contextText);
  const scopeLabel =
    scope === 'selection' ? '当前选区' : scope === 'section' ? '当前章节' : '整篇文档';

  if (/(总结|概括|摘要|梳理)/.test(normalizedPrompt)) {
    return `${scopeLabel}的分析结果如下：\n\n${summary}`;
  }

  if (/(风险|问题|注意事项|排查)/.test(normalizedPrompt)) {
    const riskLines = normalizeLineEndings(contextText)
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => /(注意|风险|错误|失败|异常|排查|提示)/.test(line))
      .slice(0, 8);

    if (riskLines.length > 0) {
      return `${scopeLabel}中识别到以下风险或注意点：\n\n- ${riskLines.join('\n- ')}`;
    }

    return `${scopeLabel}中没有识别到明显的风险提示语句。建议补充“前置条件”“失败现象”“排查方式”三类信息。`;
  }

  if (/(续写|补全|补充)/.test(normalizedPrompt)) {
    return transformWithBuiltinAi('continue', contextText);
  }

  if (/(改写|润色|优化表达)/.test(normalizedPrompt)) {
    return transformWithBuiltinAi('rewrite', contextText);
  }

  return [
    `这里是对《${fileName}》中${scopeLabel}的本地 AI 分析结果：`,
    '',
    summary,
    '',
    '如果你希望继续处理，可以尝试这些指令：',
    '- 帮我总结这一段',
    '- 帮我改写成更正式的技术文档语气',
    '- 帮我续写下一段',
    '- 帮我提炼风险点'
  ].join('\n');
}

function getOpenAiCompatibleEndpoint(apiBase: string) {
  return `${apiBase.replace(/\/$/, '')}/chat/completions`;
}

function sanitizeAiTextOutput(text: string) {
  return text
    .replace(/<(think|reasoning|analysis)>[\s\S]*?(?:<\/\1>|$)/gi, '')
    .replace(/^\s*```(?:markdown|md)?\s*/i, '')
    .replace(/\s*```\s*$/i, '')
    .trim();
}

function sanitizeAiInsertedMarkdown(text: string) {
  const cleanedLines = sanitizeAiTextOutput(text)
    .split('\n')
    .filter((line, index, all) => {
      const trimmed = line.trim();
      if (trimmed === '---') {
        return false;
      }

      if (index >= all.length - 2 && /^(如需|如果需要|若需|需要我继续|如要继续|请告诉我).*/.test(trimmed)) {
        return false;
      }

      return true;
    });

  return cleanedLines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

async function requestOpenAiCompatibleCompletion(
  settings: AISettings,
  systemPrompt: string,
  userPrompt: string
) {
  const endpoint = getOpenAiCompatibleEndpoint(settings.apiBase);
  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), aiRequestTimeoutMs);

  let response: Response;
  try {
    response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${settings.apiKey}`
      },
      body: JSON.stringify({
        model: settings.model,
        temperature: settings.temperature,
        max_tokens: settings.maxTokens,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
      }),
      signal: controller.signal
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      throw new Error(`AI 请求超时（>${Math.round(aiRequestTimeoutMs / 60000)} 分钟）`);
    }

    throw error;
  } finally {
    window.clearTimeout(timeoutId);
  }

  if (!response.ok) {
    const responseText = (await response.text()).slice(0, 400);
    throw new Error(`AI 接口返回 ${response.status}${responseText ? `：${responseText}` : ''}`);
  }

  const data = (await response.json()) as {
    choices?: Array<{ message?: { content?: string } }>;
  };
  const content = sanitizeAiTextOutput(data.choices?.[0]?.message?.content?.trim() ?? '');
  if (!content) {
    throw new Error('AI 接口没有返回内容');
  }

  return content;
}

async function testOpenAiCompatibleConnection(settings: AISettings) {
  return requestOpenAiCompatibleCompletion(
    settings,
    'You are a connectivity check endpoint. Reply with exactly: CONNECTION_OK',
    'Return exactly: CONNECTION_OK'
  );
}

function inferHeadingLevelFromPattern(text: string, hasPrimaryTitle: boolean) {
  const normalized = text.trim();
  const baseLevel = hasPrimaryTitle ? 1 : 0;

  if (/^第[一二三四五六七八九十百千万0-9]+(章|节|部分|篇)/.test(normalized)) {
    return Math.min(baseLevel + 1, 6);
  }

  if (/^[一二三四五六七八九十]+[、.．]\s*/.test(normalized)) {
    return Math.min(baseLevel + 1, 6);
  }

  if (/^(前言|引言|摘要|概述|背景|结论|总结|附录|常见问题|FAQ|部署|使用方法|注意事项|问题背景|问题分析|解决方案|安装步骤|快速开始)\b/.test(normalized)) {
    return Math.min(baseLevel + 1, 6);
  }

  return null;
}

function inferPlainTextHeadingLevel(
  line: string,
  index: number,
  lines: string[],
  hasPrimaryTitle: boolean
) {
  const text = line.trim();
  if (!text || isSpecialBlockLine(text) || isMinorStepHeading(text)) {
    return null;
  }

  const previousBlank = index === 0 || isBlankLine(lines[index - 1] ?? '');
  const nextLine = lines[index + 1] ?? '';
  const nextBlank = isBlankLine(nextLine);
  const nextLooksContent = nextLine.trim() !== '' && !isSpecialBlockLine(nextLine);
  const patternLevel = inferHeadingLevelFromPattern(text, hasPrimaryTitle);

  if (patternLevel && previousBlank) {
    return patternLevel;
  }

  if (!previousBlank || isLikelySentence(text)) {
    return null;
  }

  if (index === 0 && text.length <= 24 && isFrameworkHeading(text)) {
    return hasPrimaryTitle ? 2 : 1;
  }

  if ((nextBlank || nextLooksContent) && text.length <= 18 && isFrameworkHeading(text)) {
    return hasPrimaryTitle ? 2 : 1;
  }

  return null;
}

function inferSemanticHeadingLevel(
  text: string,
  currentLevel: number | null,
  index: number,
  lines: string[],
  hasPrimaryTitle: boolean
) {
  if (currentLevel === 1) {
    return 1;
  }

  const patternLevel = inferHeadingLevelFromPattern(text, hasPrimaryTitle);
  if (patternLevel) {
    return patternLevel;
  }

  if (currentLevel) {
    return currentLevel;
  }

  return inferPlainTextHeadingLevel(text, index, lines, hasPrimaryTitle);
}

function upsertTableOfContents(lines: string[], tocStrength: number) {
  const headings = extractHeadings(lines);
  const tocLines = buildTableOfContents(headings, tocStrength);

  if (!tocLines) {
    return { lines, changed: false, action: 'none' as const };
  }

  const existingStart = lines.findIndex((line) => /^##\s+(目录|目次|contents)$/i.test(line.trim()));

  if (existingStart >= 0) {
    let existingEnd = existingStart + 1;
    while (existingEnd < lines.length) {
      const current = lines[existingEnd];
      if (existingEnd > existingStart + 1 && /^#{1,6}\s+/.test(current)) {
        break;
      }
      existingEnd += 1;
    }

    const nextLines = [...lines];
    nextLines.splice(existingStart, existingEnd - existingStart, ...tocLines);
    return { lines: nextLines, changed: true, action: 'updated' as const };
  }

  const nextLines = [...lines];
  const firstHeadingIndex = nextLines.findIndex((line) => /^#\s+/.test(line));
  const insertIndex = firstHeadingIndex >= 0 ? Math.min(firstHeadingIndex + 2, nextLines.length) : 0;
  nextLines.splice(insertIndex, 0, ...tocLines);
  return { lines: nextLines, changed: true, action: 'inserted' as const };
}

function beautifyDocumentContent(content: string, strength: BeautifyStrength): BeautifyResult {
  const habits = analyzeWritingHabits(content);
  const changes: string[] = [];
  const input = normalizeLineEndings(content);
  const lines = input.split('\n');
  const output: string[] = [];
  let index = 0;

  while (index < lines.length) {
    const currentLine = lines[index].replace(/\s+$/g, '');
    const codeFenceMatch = currentLine.match(/^```([a-zA-Z0-9#+._-]*)?\s*$/);

    if (codeFenceMatch) {
      const codeLines: string[] = [];
      const originalLanguage = codeFenceMatch[1] || '';
      const explicitLanguage = normalizeExplicitLanguage(originalLanguage);
      let innerIndex = index + 1;

      while (innerIndex < lines.length && !/^```/.test(lines[innerIndex])) {
        codeLines.push(lines[innerIndex].replace(/\s+$/g, ''));
        innerIndex += 1;
      }

      const inferredLanguage = explicitLanguage || inferCodeLanguage(codeLines.join('\n'));
      const openingLanguage = explicitLanguage || (inferredLanguage !== 'text' ? inferredLanguage : '');

      if (!explicitLanguage && inferredLanguage !== 'text') {
        changes.push(`补全了代码块语言：${inferredLanguage}`);
      }

      if (explicitLanguage && explicitLanguage !== originalLanguage) {
        changes.push(`统一了代码块语言名称：${explicitLanguage}`);
      }

      if (strength === 'deep' && output.length > 0 && output[output.length - 1] !== '') {
        output.push('');
      }
      output.push(openingLanguage ? `\`\`\`${openingLanguage}` : '```');
      output.push(...codeLines);
      output.push('```');
      if (strength === 'deep' && lines[innerIndex + 1] && lines[innerIndex + 1].trim() !== '') {
        output.push('');
      }
      index = innerIndex + 1;
      continue;
    }

    const headingMatch = currentLine.match(/^(#{1,6})\s*(.+)$/);
    if (headingMatch) {
      if (output.length > 0 && output[output.length - 1] !== '') {
        output.push('');
      }

      output.push(`${headingMatch[1]} ${headingMatch[2].trim()}`);

      if (lines[index + 1] !== '') {
        output.push('');
      }

      index += 1;
      continue;
    }

    const bulletMatch = currentLine.match(/^(\s*)[-*+]\s+(.+)$/);
    if (bulletMatch) {
      if (strength === 'deep' && output.length > 0 && output[output.length - 1] !== '' && !/^\s*[-*+]\s+/.test(output[output.length - 1])) {
        output.push('');
      }
      output.push(`${bulletMatch[1]}${habits.bulletPreference} ${bulletMatch[2].trim()}`);
      index += 1;
      continue;
    }

    output.push(currentLine);
    index += 1;
  }

  let compacted = output.join('\n').replace(/\n{3,}/g, '\n\n').trim();

  if (strength === 'deep') {
    compacted = compacted
      .replace(/([^\n])\n(\d+\.\s+)/g, '$1\n\n$2')
      .replace(/([^\n])\n([-*+]\s+)/g, '$1\n\n$2')
      .replace(/([^\n])\n(```)/g, '$1\n\n$2')
      .replace(/(```)\n([^\n])/g, '$1\n\n$2')
      .replace(/\n{3,}/g, '\n\n');
  }

  const finalContent = `${compacted}\n`;

  if (changes.length === 0) {
    changes.push('检查了空行、列表和代码块，当前文档已经比较规整');
  }

  if (strength === 'light') {
    changes.push('使用轻度美化，保留了更多原始排版');
  }

  if (strength === 'deep') {
    changes.push('使用深度美化，增强了结构留白与段落节奏');
  }

  return {
    content: finalContent,
    summary: changes.slice(0, 3).join('；'),
    changes: Array.from(new Set(changes))
  };
}

function App() {
  const desktopApi = window.markdownApp;
  const initialAiProfiles = loadAiProfiles();
  const browserFileInputRef = useRef<HTMLInputElement | null>(null);
  const browserImageInputRef = useRef<HTMLInputElement | null>(null);
  const previewEditorRef = useRef<HTMLDivElement | null>(null);
  const lastRenderedMarkdownRef = useRef(starterDocument);
  const lastSelectionRef = useRef<{ start: number; end: number; source: SelectionSource } | null>(null);
  const lastCaretRangeRef = useRef<{ start: number; end: number; source: SelectionSource } | null>(null);
  const lastCaretDomRangeRef = useRef<Range | null>(null);
  const lastSelectedDomRangeRef = useRef<Range | null>(null);
  const undoHistoryRef = useRef<EditorHistoryEntry[]>([]);
  const redoHistoryRef = useRef<EditorHistoryEntry[]>([]);
  const pendingEditorSelectionRef = useRef<{ start: number; end: number } | null>(null);
  const lastActiveSectionIdRef = useRef<string | null>(null);
  const [viewMode, setViewMode] = useState<ViewMode>('preview');
  const [doc, setDoc] = useState<EditorFileState>({
    filePath: null,
    displayName: '欢迎.md',
    content: starterDocument
  });
  const [lastSavedContent, setLastSavedContent] = useState(starterDocument);
  const [statusMessage, setStatusMessage] = useState('准备就绪');
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>(() => loadRecentFiles());
  const [beautifyPreview, setBeautifyPreview] = useState<BeautifyPreviewState | null>(null);
  const [beautifyStrength, setBeautifyStrength] = useState<BeautifyStrength>('standard');
  const [translationDirection, setTranslationDirection] = useState<TranslationDirection>('zh-to-en');
  const [aiProfiles, setAiProfiles] = useState<AIProfile[]>(() => initialAiProfiles);
  const [activeAiProfileId, setActiveAiProfileId] = useState<string>(() => loadActiveAiProfileId(initialAiProfiles));
  const [theme, setTheme] = useState<ThemeName>(() => loadThemePreference());
  const [translationPreview, setTranslationPreview] = useState<TranslationPreviewState | null>(null);
  const [docxImportPreview, setDocxImportPreview] = useState<DocxImportPreviewState | null>(null);
  const [isDocxImporting, setIsDocxImporting] = useState(false);
  const [isTranslating, setIsTranslating] = useState(false);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [aiConfigFilePath, setAiConfigFilePath] = useState<string | null>(null);
  const [hasLoadedDesktopAiConfig, setHasLoadedDesktopAiConfig] = useState<boolean>(() => !desktopApi);
  const [isAiDrawerOpen, setIsAiDrawerOpen] = useState(false);
  const [aiContextScope, setAiContextScope] = useState<AIContextScope>('section');
  const [liveAiSectionId, setLiveAiSectionId] = useState<string | null>(null);
  const [aiMessages, setAiMessages] = useState<AIChatMessage[]>([
    {
      id: 'ai-welcome',
      role: 'assistant',
      content:
        '我是墨笺内置的文档助手。你可以围绕当前选区、当前章节或整篇文档向我提问，也可以让我帮你改写、扩写或续写内容。',
      contextScope: 'document',
      providerLabel: '内置文档助手'
    }
  ]);
  const [aiInput, setAiInput] = useState('');
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [aiActionPreview, setAiActionPreview] = useState<AIActionPreviewState | null>(null);
  const [isAiConfigOpen, setIsAiConfigOpen] = useState(false);
  const [isAiTesting, setIsAiTesting] = useState(false);
  const [aiConnectionMessage, setAiConnectionMessage] = useState('尚未测试连接');
  const [aiRequestLogs, setAiRequestLogs] = useState<AIRequestLog[]>([]);
  const [aiToastMessage, setAiToastMessage] = useState<string | null>(null);

  const isDirty = doc.content !== lastSavedContent;
  const fileName = formatFileName(doc.filePath, doc.displayName);
  const activeAiProfile = useMemo(
    () => aiProfiles.find((profile) => profile.id === activeAiProfileId) ?? aiProfiles[0] ?? createAiProfile('内置文档助手'),
    [activeAiProfileId, aiProfiles]
  );
  const aiSettings = useMemo<AISettings>(
    () => ({
      provider: activeAiProfile.provider,
      apiBase: activeAiProfile.apiBase,
      apiKey: activeAiProfile.apiKey,
      model: activeAiProfile.model,
      temperature: activeAiProfile.temperature,
      maxTokens: activeAiProfile.maxTokens
    }),
    [activeAiProfile]
  );

  const stats = useMemo(() => {
    const trimmed = doc.content.trim();
    const words = trimmed ? trimmed.split(/\s+/).length : 0;
    const characters = doc.content.length;
    const lines = normalizeLineEndings(doc.content).split('\n').length;

    return { words, characters, lines };
  }, [doc.content]);

  const outline = useMemo<OutlineItem[]>(() => extractOutlineFromMarkdown(doc.content), [doc.content]);
  const editableHtml = useMemo(() => buildEditableHtml(doc.content, outline), [doc.content, outline]);
  const effectiveAiSectionId = liveAiSectionId ?? lastActiveSectionIdRef.current;
  const effectiveAiSection = getSectionRangeByOutlineId(effectiveAiSectionId);
  const shouldHighlightAiSection = isAiDrawerOpen && aiContextScope === 'section';

  useEffect(() => {
    if (outline.length === 0) {
      lastActiveSectionIdRef.current = null;
      setLiveAiSectionId(null);
      return;
    }

    const currentId = lastActiveSectionIdRef.current;
    if (!currentId || !outline.some((item) => item.id === currentId)) {
      lastActiveSectionIdRef.current = outline[0].id;
    }

    setLiveAiSectionId((current) => (current && outline.some((item) => item.id === current) ? current : outline[0].id));
  }, [outline]);

  const beautifyDiffRows = useMemo(() => {
    if (!beautifyPreview) {
      return [];
    }

    return buildDiffRows(beautifyPreview.original, beautifyPreview.beautified);
  }, [beautifyPreview]);

  const translationDiffRows = useMemo(() => {
    if (!translationPreview) {
      return [];
    }

    return buildDiffRows(translationPreview.original, translationPreview.translated);
  }, [translationPreview]);

  const aiActionDiffRows = useMemo(() => {
    if (!aiActionPreview) {
      return [];
    }

    return buildDiffRows(aiActionPreview.original, aiActionPreview.result);
  }, [aiActionPreview]);

  const docxImportPreviewHtml = useMemo(() => {
    if (!docxImportPreview) {
      return '';
    }

    return sanitizeHtml(docxImportPreview.previewHtml);
  }, [docxImportPreview]);

  const canInsertDocxImportIntoCurrentDocument =
    !docxImportPreview || docxImportPreview.imageAssets.length === 0 || Boolean(doc.filePath);

  useEffect(() => {
    if (!desktopApi) {
      setStatusMessage('当前为浏览器预览模式');
      return;
    }

    setStatusMessage(`桌面应用已连接（${desktopApi.ping()}）`);
  }, [desktopApi]);

  useEffect(() => {
    document.title = `${isDirty ? '* ' : ''}${fileName} - 墨笺 Markdown 编辑器`;
  }, [fileName, isDirty]);

  useEffect(() => {
    if (!desktopApi) {
      return;
    }

    let cancelled = false;

    void desktopApi
      .readAiConfigDocument()
      .then((result) => {
        if (cancelled) {
          return;
        }

        setAiConfigFilePath(result.filePath);

        if (result.profiles.length > 0) {
          const loadedProfiles = buildAiProfilesFromPersistedConfig(result.profiles);
          const nextActiveProfile =
            loadedProfiles.find((profile) => profile.name === result.activeProfileName) ?? loadedProfiles[0] ?? null;

          setAiProfiles(loadedProfiles);
          setActiveAiProfileId(nextActiveProfile?.id ?? '');
          setAiConnectionMessage(`已从配置文件加载 ${loadedProfiles.length} 个模型配置。`);
        }
      })
      .catch((error) => {
        console.error('[renderer] load ai config file failed', error);
        if (!cancelled) {
          setAiConnectionMessage('配置文件读取失败，已回退到本地配置。');
        }
      })
      .finally(() => {
        if (!cancelled) {
          setHasLoadedDesktopAiConfig(true);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [desktopApi]);

  useEffect(() => {
    const root = previewEditorRef.current;
    if (!root) {
      return;
    }

    if (document.activeElement === root && lastRenderedMarkdownRef.current === doc.content) {
      return;
    }

    root.innerHTML = editableHtml;
    lastRenderedMarkdownRef.current = doc.content;

    if (pendingEditorSelectionRef.current) {
      setContentEditableSelection(
        root,
        pendingEditorSelectionRef.current.start,
        pendingEditorSelectionRef.current.end
      );
      root.focus();
      pendingEditorSelectionRef.current = null;
    }
  }, [doc.content, editableHtml]);

  useEffect(() => {
    const root = previewEditorRef.current;
    if (!root) {
      return;
    }

    const activeSectionId = shouldHighlightAiSection ? effectiveAiSectionId : null;
    const children = Array.from(root.children) as HTMLElement[];

    children.forEach((element) => {
      const outlineId = element.dataset.outlineId ?? null;
      const isActiveHeading = Boolean(activeSectionId && outlineId === activeSectionId);
      element.classList.toggle('is-ai-section-active', isActiveHeading);
      element.classList.toggle('is-ai-section-heading', isActiveHeading);
    });
  }, [doc.content, editableHtml, effectiveAiSectionId, shouldHighlightAiSection]);

  useEffect(() => {
    const handleSelectionChange = () => {
      const root = previewEditorRef.current;
      const selection = window.getSelection();

      if (!root || !selection || selection.rangeCount === 0) {
        return;
      }

      const range = selection.getRangeAt(0);
      if (!root.contains(range.commonAncestorContainer)) {
        return;
      }

      updateEditorSelectionState('preview');
    };

    document.addEventListener('selectionchange', handleSelectionChange);
    return () => {
      document.removeEventListener('selectionchange', handleSelectionChange);
    };
  }, [doc.content]);

  useEffect(() => {
    if (!isAiDrawerOpen || aiContextScope !== 'section') {
      return;
    }

    const root = previewEditorRef.current;
    const syncSection = () => {
      const focusedEditor = Boolean(root && document.activeElement === root);
      const section =
        (focusedEditor ? getCurrentCaretSectionRange() : null) ??
        (root ? getSectionRangeFromVisualOffset(root.scrollTop + 48) : null) ??
        getCurrentVisibleSectionRange();

      if (section?.id) {
        lastActiveSectionIdRef.current = section.id;
        setLiveAiSectionId((current) => (current === section.id ? current : section.id));
      }
    };

    syncSection();
    if (!root) {
      return;
    }

    root.addEventListener('scroll', syncSection, { passive: true });
    return () => {
      root.removeEventListener('scroll', syncSection);
    };
  }, [aiContextScope, isAiDrawerOpen, doc.content]);

  useEffect(() => {
    window.localStorage.setItem(recentFilesStorageKey, JSON.stringify(recentFiles));
  }, [recentFiles]);

  useEffect(() => {
    const normalizedProfiles = normalizeAiProfilesForPersistence(aiProfiles);
    window.localStorage.setItem(aiProfilesStorageKey, JSON.stringify(toLocalSafeAiProfiles(normalizedProfiles)));
    window.localStorage.setItem(aiSettingsStorageKey, JSON.stringify({ ...aiSettings, apiKey: '' }));

    if (!desktopApi || !hasLoadedDesktopAiConfig) {
      return;
    }

    const activeProfileName =
      normalizedProfiles.find((profile) => profile.id === activeAiProfileId)?.name ?? normalizedProfiles[0]?.name ?? null;

    void desktopApi
      .writeAiConfigDocument({
        activeProfileName,
        profiles: toPersistedAiProfiles(normalizedProfiles)
      })
      .then((result) => {
        setAiConfigFilePath(result.filePath);
      })
      .catch((error) => {
        console.error('[renderer] persist ai config file failed', error);
      });
  }, [activeAiProfileId, aiProfiles, aiSettings, desktopApi, hasLoadedDesktopAiConfig]);

  useEffect(() => {
    if (!activeAiProfileId || !aiProfiles.some((profile) => profile.id === activeAiProfileId)) {
      const fallbackId = aiProfiles[0]?.id ?? '';
      if (fallbackId && fallbackId !== activeAiProfileId) {
        setActiveAiProfileId(fallbackId);
      }
      return;
    }

    window.localStorage.setItem(aiActiveProfileStorageKey, activeAiProfileId);
  }, [activeAiProfileId, aiProfiles]);

  useEffect(() => {
    window.localStorage.setItem(themeStorageKey, theme);
  }, [theme]);

  useEffect(() => {
    document.body.dataset.theme = theme;
    return () => {
      delete document.body.dataset.theme;
    };
  }, [theme]);

  useEffect(() => {
    if (!aiToastMessage) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setAiToastMessage(null);
    }, 5000);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [aiToastMessage]);

  useEffect(() => {
    if (!desktopApi) {
      return;
    }

    desktopApi.syncDocumentState({
      filePath: doc.filePath,
      displayName: doc.displayName,
      content: doc.content,
      isDirty
    });
  }, [desktopApi, doc.content, doc.displayName, doc.filePath, isDirty]);

  useEffect(() => {
    if (!desktopApi) {
      return;
    }

    const dispose = desktopApi.onMenuAction(async (action) => {
      if (action === 'new-file') {
        handleNewDocument();
      }

      if (action === 'open-file') {
        await handleOpenFile();
      }

      if (action === 'import-docx') {
        await handleImportDocx();
      }

      if (action === 'save-file') {
        await handleSaveFile(false);
      }

      if (action === 'save-file-as') {
        await handleSaveFile(true);
      }

      if (action === 'reference-image') {
        await handleReferenceImage();
      }

      if (action === 'open-ai-chat') {
        setIsAiDrawerOpen(true);
        setStatusMessage('已打开 AI 对话面板');
      }

      if (action === 'open-ai-config') {
        setIsAiConfigOpen(true);
        setStatusMessage('已打开模型配置中心');
      }

      if (action === 'ai-rewrite') {
        await handleAiAction('rewrite');
      }

      if (action === 'ai-expand') {
        await handleAiAction('expand');
      }

      if (action === 'ai-shorten') {
        await handleAiAction('shorten');
      }

      if (action === 'ai-continue') {
        await handleAiAction('continue');
      }

      if (action === 'ai-beautify') {
        openBeautifyPreview();
      }

      if (action === 'ai-translate') {
        await openTranslationPreview();
      }

      if (action === 'export-html') {
        await handleExportHtml();
      }

      if (action === 'export-pdf') {
        await handleExportPdf();
      }

      if (action === 'view-write' || action === 'view-split' || action === 'view-preview') {
        setViewMode('preview');
        setStatusMessage('当前使用统一的实时预览编辑工作区');
      }

      if (action === 'theme-paper') {
        applyTheme('paper');
      }

      if (action === 'theme-mist') {
        applyTheme('mist');
      }

      if (action === 'theme-slate') {
        applyTheme('slate');
      }

      if (action === 'theme-graphite') {
        applyTheme('graphite');
      }

      if (action === 'theme-terminal') {
        applyTheme('terminal');
      }

      if (action === 'theme-nightcode') {
        applyTheme('nightcode');
      }

      if (action === 'theme-campus') {
        applyTheme('campus');
      }

      if (action === 'theme-youth') {
        applyTheme('youth');
      }
    });

    return () => {
      dispose();
    };
  }, [desktopApi, doc.content, doc.filePath, isDirty, aiContextScope, aiSettings, fileName]);

  useEffect(() => {
    if (!contextMenu) {
      return;
    }

    const handleClose = () => {
      setContextMenu(null);
    };

    const handleEscape = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        setContextMenu(null);
      }
    };

    window.addEventListener('click', handleClose);
    window.addEventListener('resize', handleClose);
    window.addEventListener('scroll', handleClose, true);
    window.addEventListener('keydown', handleEscape);

    return () => {
      window.removeEventListener('click', handleClose);
      window.removeEventListener('resize', handleClose);
      window.removeEventListener('scroll', handleClose, true);
      window.removeEventListener('keydown', handleEscape);
    };
  }, [contextMenu]);

  useEffect(() => {
    const handleGlobalUndoRedo = (event: KeyboardEvent) => {
      if (!(event.ctrlKey || event.metaKey) || event.altKey) {
        return;
      }

      const target = event.target as HTMLElement | null;
      const tagName = target?.tagName?.toLowerCase();
      const isEditableTarget =
        Boolean(target?.closest('.preview-editor--workspace')) ||
        target?.isContentEditable ||
        tagName === 'input' ||
        tagName === 'textarea' ||
        tagName === 'select';

      if (isEditableTarget) {
        return;
      }

      if (handleUndoRedoShortcut(event.key, event.shiftKey)) {
        event.preventDefault();
        event.stopPropagation();
      }
    };

    window.addEventListener('keydown', handleGlobalUndoRedo);
    return () => {
      window.removeEventListener('keydown', handleGlobalUndoRedo);
    };
  }, [doc.content]);

  function confirmDiscardChanges() {
    if (!isDirty) {
      return true;
    }

    return window.confirm('当前内容尚未保存，是否放弃这些修改？');
  }

  function focusEditorSoon(source: SelectionSource = 'preview', selection?: number | { start: number; end: number }) {
    window.setTimeout(() => {
      const target = previewEditorRef.current;
      target?.focus();
      if (!target) {
        return;
      }

      if (typeof selection === 'number') {
        setContentEditableSelection(target, selection, selection);
      } else if (selection) {
        setContentEditableSelection(target, selection.start, selection.end);
      }
    }, 0);
  }

  function rememberRecentFile(filePath: string | null, displayName: string) {
    if (!filePath) {
      return;
    }

    setRecentFiles((current) => {
      const next = current.filter((item) => item.filePath !== filePath);
      next.unshift({ filePath, displayName });
      return next.slice(0, 8);
    });
  }

  function getSectionRangeByOutlineId(outlineId: string | null) {
    if (!outlineId) {
      return null;
    }

    const index = outline.findIndex((item) => item.id === outlineId);
    if (index < 0) {
      return null;
    }

    const current = outline[index];
    const next = outline[index + 1];
    return {
      id: current.id,
      scope: 'section' as const,
      start: current.start,
      end: next?.start ?? doc.content.length,
      text: doc.content.slice(current.start, next?.start ?? doc.content.length).trim()
    };
  }

  function getPreferredAiSectionRange() {
    return getSectionRangeByOutlineId(liveAiSectionId) ?? getCurrentVisibleSectionRange();
  }

  function getOutlineItemById(outlineId: string | null) {
    if (!outlineId) {
      return null;
    }

    return outline.find((item) => item.id === outlineId) ?? null;
  }

  function normalizeMarkdownHeadingLevels(markdown: string, parentHeadingLevel?: number | null) {
    if (!parentHeadingLevel || parentHeadingLevel >= 6) {
      return markdown;
    }

    const lines = normalizeLineEndings(markdown).split('\n');
    let inFence = false;
    let fenceMarker = '';
    let minHeadingLevel: number | null = null;

    lines.forEach((line) => {
      const fenceMatch = line.match(/^(\s*)(`{3,}|~{3,})/);
      if (fenceMatch) {
        const marker = fenceMatch[2][0];
        if (!inFence) {
          inFence = true;
          fenceMarker = marker;
        } else if (fenceMarker === marker) {
          inFence = false;
          fenceMarker = '';
        }
        return;
      }

      if (inFence) {
        return;
      }

      const headingMatch = line.match(/^(#{1,6})(\s+.*)$/);
      if (!headingMatch) {
        return;
      }

      const level = headingMatch[1].length;
      minHeadingLevel = minHeadingLevel === null ? level : Math.min(minHeadingLevel, level);
    });

    if (minHeadingLevel === null || minHeadingLevel > parentHeadingLevel) {
      return markdown;
    }

    const delta = parentHeadingLevel + 1 - minHeadingLevel;
    inFence = false;
    fenceMarker = '';

    return lines
      .map((line) => {
        const fenceMatch = line.match(/^(\s*)(`{3,}|~{3,})/);
        if (fenceMatch) {
          const marker = fenceMatch[2][0];
          if (!inFence) {
            inFence = true;
            fenceMarker = marker;
          } else if (fenceMarker === marker) {
            inFence = false;
            fenceMarker = '';
          }
          return line;
        }

        if (inFence) {
          return line;
        }

        const headingMatch = line.match(/^(#{1,6})(\s+.*)$/);
        if (!headingMatch) {
          return line;
        }

        const nextLevel = Math.min(6, headingMatch[1].length + delta);
        return `${'#'.repeat(nextLevel)}${headingMatch[2]}`;
      })
      .join('\n');
  }

  function prepareAiMarkdownForSection(markdown: string, sectionId?: string | null) {
    const parentHeadingLevel = getOutlineItemById(sectionId ?? null)?.level ?? null;
    return normalizeMarkdownHeadingLevels(sanitizeAiInsertedMarkdown(markdown), parentHeadingLevel);
  }

  function getSectionRangeFromVisualOffset(visualOffset: number) {
    const root = previewEditorRef.current;
    if (!root || outline.length === 0) {
      return null;
    }

    let activeId = outline[0]?.id ?? null;

    outline.forEach((item, index) => {
      const target = root.querySelector<HTMLElement>(`[data-outline-id="${item.id}"]`);
      if (!target) {
        return;
      }

      const next = outline[index + 1];
      const nextTarget = next ? root.querySelector<HTMLElement>(`[data-outline-id="${next.id}"]`) : null;
      const startY = target.offsetTop;
      const endY = nextTarget?.offsetTop ?? root.scrollHeight + 1;

      if (visualOffset >= startY && visualOffset < endY) {
        activeId = item.id;
      }
    });

    return getSectionRangeByOutlineId(activeId);
  }

  function getSectionRangeFromBlockElement(block: HTMLElement | null) {
    const root = previewEditorRef.current;
    if (!root || !block) {
      return null;
    }

    const rootLevelBlock = getRootLevelEditorBlock(root, block);
    if (!rootLevelBlock) {
      return null;
    }

    const directSectionId = rootLevelBlock.dataset.sectionId ?? rootLevelBlock.dataset.outlineId ?? null;
    if (directSectionId) {
      return getSectionRangeByOutlineId(directSectionId);
    }

    let sibling: HTMLElement | null = rootLevelBlock;
    while (sibling) {
      const sectionId = sibling.dataset.sectionId ?? sibling.dataset.outlineId ?? null;
      if (sectionId) {
        return getSectionRangeByOutlineId(sectionId);
      }
      sibling = sibling.previousElementSibling as HTMLElement | null;
    }

    sibling = rootLevelBlock.nextElementSibling as HTMLElement | null;
    while (sibling) {
      const sectionId = sibling.dataset.sectionId ?? sibling.dataset.outlineId ?? null;
      if (sectionId) {
        return getSectionRangeByOutlineId(sectionId);
      }
      sibling = sibling.nextElementSibling as HTMLElement | null;
    }

    const visualMiddle = rootLevelBlock.offsetTop + Math.max(1, rootLevelBlock.offsetHeight / 2);
    return getSectionRangeFromVisualOffset(visualMiddle);
  }

  function getCurrentCaretSectionRange() {
    const root = previewEditorRef.current;
    const selection = window.getSelection();

    if (!root || !selection || selection.rangeCount === 0) {
      return null;
    }

    const range = selection.getRangeAt(0);
    if (!root.contains(range.commonAncestorContainer)) {
      return null;
    }

    const blockSection = getSectionRangeFromBlockElement(getClosestEditorBlock(root, range.startContainer));
    if (blockSection) {
      return blockSection;
    }

    const rootRect = root.getBoundingClientRect();
    const measureRange = range.cloneRange();
    measureRange.collapse(true);
    const rect = measureRange.getClientRects()[0] ?? measureRange.getBoundingClientRect();
    if (!rect) {
      return null;
    }

    const caretMiddle = rect.top - rootRect.top + root.scrollTop + Math.max(1, rect.height / 2);
    return getSectionRangeFromVisualOffset(caretMiddle);
  }

  function getSectionRangeFromClientPoint(clientX: number, clientY: number) {
    const root = previewEditorRef.current;
    if (!root || outline.length === 0) {
      return null;
    }

    const pointRange = resolveRangeFromPoint(clientX, clientY);
    const fromPoint = getSectionRangeFromBlockElement(getClosestEditorBlock(root, pointRange?.startContainer ?? null));
    if (fromPoint) {
      return fromPoint;
    }

    const rootRect = root.getBoundingClientRect();
    const yWithinEditor = clientY - rootRect.top + root.scrollTop + 4;
    return getSectionRangeFromVisualOffset(yWithinEditor);
  }

  function getCurrentVisibleSectionRange() {
    if (outline.length === 0) {
      const text = doc.content.trim();
      return text
        ? {
            scope: 'document' as const,
            start: 0,
            end: doc.content.length,
            text
          }
        : null;
    }

    const root = previewEditorRef.current;
    const isEditorFocused = Boolean(root && document.activeElement === root);
    const caretSection = isEditorFocused ? getCurrentCaretSectionRange() : null;
    if (caretSection) {
      lastActiveSectionIdRef.current = caretSection.id;
      return caretSection;
    }

    if (!root) {
      const rememberedSection = getSectionRangeByOutlineId(lastActiveSectionIdRef.current);
      if (rememberedSection) {
        return rememberedSection;
      }

      const firstSection = getSectionRangeByOutlineId(outline[0]?.id ?? null);
      if (firstSection) {
        lastActiveSectionIdRef.current = firstSection.id;
      }
      return firstSection;
    }

    const viewportTop = root.scrollTop + 48;
    let activeId = outline[0]?.id ?? null;

    outline.forEach((item) => {
      const target = root.querySelector<HTMLElement>(`[data-outline-id="${item.id}"]`);
      if (target && target.offsetTop <= viewportTop) {
        activeId = item.id;
      }
    });

    lastActiveSectionIdRef.current = activeId;
    const visibleSection = getSectionRangeByOutlineId(activeId);
    if (visibleSection) {
      return visibleSection;
    }

    return getSectionRangeByOutlineId(lastActiveSectionIdRef.current);
  }

  function locateSelectedTextInDocument(selectedText: string) {
    const text = selectedText.trim();
    if (!text) {
      return null;
    }

    const preferredSection = getCurrentVisibleSectionRange();
    if (preferredSection) {
      const sectionIndex = doc.content.indexOf(text, preferredSection.start);
      if (sectionIndex >= 0 && sectionIndex < preferredSection.end) {
        return {
          start: sectionIndex,
          end: sectionIndex + text.length,
          text,
          source: 'preview' as const
        };
      }
    }

    const globalIndex = doc.content.indexOf(text);
    if (globalIndex >= 0) {
      return {
        start: globalIndex,
        end: globalIndex + text.length,
        text,
        source: 'preview' as const
      };
    }

    return null;
  }

  function getPreferredSelection() {
    const liveSelection = locateSelectedTextInDocument(getRichEditorTextSelection());
    if (liveSelection) {
      return liveSelection;
    }

    if (!lastSelectionRef.current) {
      return null;
    }

    const { start, end, source } = lastSelectionRef.current;
    if (end <= start) {
      return null;
    }

    return {
      start,
      end,
      text: doc.content.slice(start, end),
      source
    };
  }

  function resolveAiContext(scope: AIContextScope) {
    if (scope === 'selection') {
      const selection = getPreferredSelection();
      if (selection && selection.text.trim()) {
        return {
          scope,
          text: selection.text.trim(),
          selectionRange: { start: selection.start, end: selection.end },
          sectionId: null,
          sectionRange: null
        };
      }
    }

    if (scope === 'section') {
      const section = getPreferredAiSectionRange();
      if (section && section.text.trim()) {
        return {
          scope: 'section' as const,
          text: section.text.trim(),
          selectionRange: null,
          sectionId: section.id,
          sectionRange: { start: section.start, end: section.end }
        };
      }
    }

    return {
      scope: 'document' as const,
      text: doc.content.trim(),
      selectionRange: null,
      sectionId: null,
      sectionRange: null
    };
  }

  function syncRichEditorToDoc() {
    const root = previewEditorRef.current;
    if (!root) {
      return;
    }

    const markdown = serializeRichEditorContent(root);
    lastRenderedMarkdownRef.current = markdown;
    setDoc((current) => (current.content === markdown ? current : { ...current, content: markdown }));
  }

  function pushUndoHistoryEntry(entry: EditorHistoryEntry) {
    const last = undoHistoryRef.current[undoHistoryRef.current.length - 1];
    if (last && last.content === entry.content) {
      return;
    }

    undoHistoryRef.current.push(entry);
    if (undoHistoryRef.current.length > 100) {
      undoHistoryRef.current.shift();
    }
  }

  function restoreHistoryEntry(
    entry: EditorHistoryEntry | undefined,
    targetStack: { current: EditorHistoryEntry[] },
    status: string
  ) {
    if (!entry) {
      return;
    }

    const currentSelection = previewEditorRef.current ? getContentEditableSelectionOffsets(previewEditorRef.current) : null;

    targetStack.current.push({
      content: doc.content,
      selection: currentSelection
    });
    if (targetStack.current.length > 100) {
      targetStack.current.shift();
    }

    pendingEditorSelectionRef.current = entry.selection;
    lastRenderedMarkdownRef.current = '';
    setDoc((current) => ({
      ...current,
      content: entry.content
    }));
    setStatusMessage(status);
  }

  function handleEditorUndo() {
    const entry = undoHistoryRef.current.pop();
    restoreHistoryEntry(entry, redoHistoryRef, '已撤销');
  }

  function handleEditorRedo() {
    const entry = redoHistoryRef.current.pop();
    restoreHistoryEntry(entry, undoHistoryRef, '已重做');
  }

  function handleUndoRedoShortcut(key: string, shiftKey: boolean) {
    const normalizedKey = key.toLowerCase();

    if (normalizedKey === 'z' && !shiftKey) {
      handleEditorUndo();
      return true;
    }

    if (normalizedKey === 'y' || (normalizedKey === 'z' && shiftKey)) {
      handleEditorRedo();
      return true;
    }

    return false;
  }

  function recordProgrammaticDocumentChange() {
    pushUndoHistoryEntry({
      content: doc.content,
      selection: previewEditorRef.current ? getContentEditableSelectionOffsets(previewEditorRef.current) : null
    });
    redoHistoryRef.current = [];
  }

  function resetUndoHistory() {
    undoHistoryRef.current = [];
    redoHistoryRef.current = [];
  }

  function getRichEditorTextSelection() {
    const root = previewEditorRef.current;
    const selection = window.getSelection();

    if (!root || !selection || selection.rangeCount === 0) {
      return '';
    }

    const range = selection.getRangeAt(0);
    if (!root.contains(range.commonAncestorContainer)) {
      return '';
    }

    return selection.toString().trim();
  }

  function restoreSavedEditorRange(kind: 'caret' | 'selection') {
    const root = previewEditorRef.current;
    const selection = window.getSelection();
    const savedRange = kind === 'selection' ? lastSelectedDomRangeRef.current : lastCaretDomRangeRef.current;

    if (!root || !selection || !savedRange) {
      return null;
    }

    if (!root.contains(savedRange.commonAncestorContainer)) {
      return null;
    }

    const nextRange = savedRange.cloneRange();
    selection.removeAllRanges();
    selection.addRange(nextRange);
    return nextRange;
  }

  function updateEditorSelectionState(source: SelectionSource) {
    const root = previewEditorRef.current;
    const selection = window.getSelection();
    if (!root || !selection || selection.rangeCount === 0) {
      return;
    }

    const offsets = getContentEditableSelectionOffsets(root);
    lastCaretRangeRef.current = offsets ? { ...offsets, source } : null;
    const liveRange = selection.getRangeAt(0);
    if (root.contains(liveRange.commonAncestorContainer)) {
      const caretRange = liveRange.cloneRange();
      caretRange.collapse(false);
      lastCaretDomRangeRef.current = caretRange;
      lastSelectedDomRangeRef.current = liveRange.collapsed ? null : liveRange.cloneRange();
    }

    const selected = locateSelectedTextInDocument(getRichEditorTextSelection());
    lastSelectionRef.current = selected ? { start: selected.start, end: selected.end, source } : null;

    const activeSection = getCurrentCaretSectionRange();
    if (activeSection?.id) {
      lastActiveSectionIdRef.current = activeSection.id;
      setLiveAiSectionId(activeSection.id);
    }
  }

  function insertMarkdownIntoRichEditor(
    markdown: string,
    rangeOverride?: Range | null,
    options?: { recordHistory?: boolean }
  ) {
    const root = previewEditorRef.current;
    const selection = window.getSelection();

    if (!root || !selection || selection.rangeCount === 0) {
      return false;
    }

    const range = rangeOverride ? rangeOverride.cloneRange() : selection.getRangeAt(0);
    if (!root.contains(range.commonAncestorContainer)) {
      return false;
    }

    if (rangeOverride) {
      selection.removeAllRanges();
      selection.addRange(range);
    }

    if (options?.recordHistory) {
      pushUndoHistoryEntry({
        content: doc.content,
        selection: getContentEditableSelectionOffsets(root)
      });
      redoHistoryRef.current = [];
    }

    const fragmentMarkup = sanitizeHtml(renderMarkdown(markdown, false));
    const template = document.createElement('template');
    template.innerHTML = fragmentMarkup;
    const fragment = template.content.cloneNode(true) as DocumentFragment;
    const lastNode = fragment.lastChild;

    range.deleteContents();
    range.insertNode(fragment);

    if (lastNode) {
      const nextRange = document.createRange();
      nextRange.selectNodeContents(lastNode);
      nextRange.collapse(false);
      selection.removeAllRanges();
      selection.addRange(nextRange);
    }

    syncRichEditorToDoc();
    updateEditorSelectionState('preview');
    root.focus();
    return true;
  }

  function replaceContentRange(
    start: number,
    end: number,
    nextText: string,
    source: SelectionSource = 'preview',
    nextSelection?: { start: number; end: number }
  ) {
    const fallbackCursor = start + nextText.length;
    const selection = nextSelection ?? { start: fallbackCursor, end: fallbackCursor };
    const nextContent = `${doc.content.slice(0, start)}${nextText}${doc.content.slice(end)}`;

    if (nextContent === doc.content) {
      return;
    }

    recordProgrammaticDocumentChange();

    setDoc((current) => ({
      ...current,
      content: `${current.content.slice(0, start)}${nextText}${current.content.slice(end)}`
    }));

    lastSelectionRef.current = {
      start: selection.start,
      end: selection.end,
      source
    };

    focusEditorSoon(source, selection);
  }

  function insertIntoCurrentPosition(
    markdown: string,
    options?: { preferProgrammatic?: boolean; sectionId?: string | null }
  ) {
    const sanitized = options?.preferProgrammatic ? prepareAiMarkdownForSection(markdown, options.sectionId) : markdown;
    const restoredCaretRange = restoreSavedEditorRange('caret');

    if (insertMarkdownIntoRichEditor(sanitized, restoredCaretRange, { recordHistory: true })) {
      return true;
    }

    const selection = getPreferredSelection();
    const caret = lastCaretRangeRef.current;

    if (!selection) {
      if (caret) {
        replaceContentRange(caret.start, caret.end, sanitized, caret.source, {
          start: caret.start + sanitized.length,
          end: caret.start + sanitized.length
        });
        return true;
      }

      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return false;
    }

    replaceContentRange(selection.start, selection.end, sanitized, selection.source);
    return true;
  }

  function getCurrentInsertionTarget() {
    const selection = getPreferredSelection();
    if (selection) {
      return {
        start: selection.start,
        end: selection.end,
        text: selection.text,
        source: selection.source,
        hasSelection: true
      };
    }

    const caret = lastCaretRangeRef.current;
    if (caret) {
      return {
        start: caret.start,
        end: caret.end,
        text: '',
        source: caret.source,
        hasSelection: false
      };
    }

    return null;
  }

  function replaceCurrentInsertionTarget(
    target:
      | {
          start: number;
          end: number;
          text: string;
          source: SelectionSource;
          hasSelection: boolean;
        }
      | null,
    nextText: string,
    status: string,
    nextSelection?: { start: number; end: number }
  ) {
    if (!target) {
      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return false;
    }

    replaceContentRange(target.start, target.end, nextText, target.source, nextSelection && {
      start: target.start + nextSelection.start,
      end: target.start + nextSelection.end
    });
    setStatusMessage(status);
    return true;
  }

  function insertInlineMarkdown(before: string, after: string, fallbackText: string, status: string) {
    const target = getCurrentInsertionTarget();
    if (!target) {
      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return;
    }

    const body = target.hasSelection ? target.text : fallbackText;
    replaceCurrentInsertionTarget(target, `${before}${body}${after}`, status, {
      start: before.length,
      end: before.length + body.length
    });
  }

  function insertPrefixedLinesAtCurrentPosition(
    fallbackLines: string[],
    prefixer: (line: string, index: number) => string,
    status: string
  ) {
    const target = getCurrentInsertionTarget();
    if (!target) {
      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return;
    }

    const sourceLines = target.hasSelection ? normalizeLineEndings(target.text).split('\n') : fallbackLines;
    const nextText = sourceLines.map((line, index) => prefixer(line, index)).join('\n');
    replaceCurrentInsertionTarget(target, nextText, status, { start: 0, end: nextText.length });
  }

  function insertLinkAtCurrentPosition() {
    const target = getCurrentInsertionTarget();
    if (!target) {
      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return;
    }

    const label = target.hasSelection ? target.text : '链接文字';
    const url = 'https://example.com';
    const nextText = `[${label}](${url})`;
    const urlStart = label.length + 3;
    replaceCurrentInsertionTarget(target, nextText, '已插入链接', {
      start: urlStart,
      end: urlStart + url.length
    });
  }

  function insertImageSyntaxAtCurrentPosition() {
    const target = getCurrentInsertionTarget();
    if (!target) {
      setStatusMessage('无法确定插入位置，请先在正文中单击目标位置');
      return;
    }

    const alt = target.hasSelection ? target.text : '图片描述';
    const path = 'images/example.png';
    const nextText = `![${alt}](${path})`;
    const pathStart = alt.length + 4;
    replaceCurrentInsertionTarget(target, nextText, '已插入图片语法', {
      start: pathStart,
      end: pathStart + path.length
    });
  }

  function insertTableAtCurrentPosition() {
    const target = getCurrentInsertionTarget();
    const snippet = ['| 列 1 | 列 2 | 列 3 |', '| --- | --- | --- |', '| 内容 1 | 内容 2 | 内容 3 |'].join('\n');
    replaceCurrentInsertionTarget(target, snippet, '已插入表格', {
      start: 2,
      end: 5
    });
  }

  function insertDividerAtCurrentPosition() {
    const target = getCurrentInsertionTarget();
    const divider = target?.hasSelection ? `${target.text}\n\n---\n` : '\n---\n';
    replaceCurrentInsertionTarget(target, divider, '已插入分割线', target?.hasSelection ? { start: 0, end: target.text.length } : undefined);
  }

  function insertHeadingAtCurrentPosition(level: 1 | 2 | 3 | 4 | 5 | 6) {
    const prefix = `${'#'.repeat(level)} `;
    const fallbackText = `标题 ${level}`;
    const richSelectionText = getRichEditorTextSelection();

    if (insertMarkdownIntoRichEditor(`${prefix}${richSelectionText || fallbackText}`, null, { recordHistory: true })) {
      setStatusMessage(`已插入 ${fallbackText}`);
      return;
    }

    const selection = getPreferredSelection();

    if (!selection) {
      const appendPrefix = doc.content.endsWith('\n') ? '' : '\n\n';
      const snippet = `${appendPrefix}${prefix}${fallbackText}\n\n`;
      const start = doc.content.length + appendPrefix.length + prefix.length;
      const end = start + fallbackText.length;
      recordProgrammaticDocumentChange();
      setDoc((current) => ({
        ...current,
        content: `${current.content}${snippet}`
      }));
      lastSelectionRef.current = { start, end, source: 'preview' };
      focusEditorSoon('preview', { start, end });
      setStatusMessage(`已插入 ${fallbackText}`);
      return;
    }

    const selectedText = selection.text.trim();
    const titleText = selectedText || fallbackText;
    const nextText = `${prefix}${titleText}`;
    const selectStart = selection.start + prefix.length;
    const selectEnd = selectStart + titleText.length;
    replaceContentRange(selection.start, selection.end, nextText, selection.source, {
      start: selectStart,
      end: selectEnd
    });
    setStatusMessage(`已插入 ${fallbackText}`);
  }

  function insertCodeBlockAtCurrentPosition() {
    const richSelectionText = getRichEditorTextSelection();

    if (insertMarkdownIntoRichEditor(`\`\`\`\n${richSelectionText || '在此输入代码'}\n\`\`\``, null, { recordHistory: true })) {
      setStatusMessage('已插入代码块');
      return;
    }

    const selection = getPreferredSelection();

    if (!selection) {
      const appendPrefix = doc.content.endsWith('\n') ? '' : '\n\n';
      const snippet = `${appendPrefix}\`\`\`\n在此输入代码\n\`\`\`\n`;
      const start = doc.content.length + appendPrefix.length + 4;
      const end = start + '在此输入代码'.length;
      recordProgrammaticDocumentChange();
      setDoc((current) => ({
        ...current,
        content: `${current.content}${snippet}`
      }));
      lastSelectionRef.current = { start, end, source: 'preview' };
      focusEditorSoon('preview', { start, end });
      setStatusMessage('已插入代码块');
      return;
    }

    const wrapped = `\`\`\`\n${selection.text || '在此输入代码'}\n\`\`\``;
    const start = selection.start + 4;
    const end = start + (selection.text || '在此输入代码').length;
    replaceContentRange(selection.start, selection.end, wrapped, selection.source, { start, end });
    setStatusMessage('已插入代码块');
  }

  function createEmptyDocument() {
    resetUndoHistory();
    setDoc({
      filePath: null,
      displayName: '未命名.md',
      content: '# 未命名文档\n\n'
    });
    setLastSavedContent('');
    setViewMode('preview');
    setBeautifyPreview(null);
    setTranslationPreview(null);
    setDocxImportPreview(null);
    setStatusMessage('已新建文档');
    focusEditorSoon('preview');
  }

  function handleNewDocument() {
    if (!confirmDiscardChanges()) {
      return;
    }

    createEmptyDocument();
  }

  async function openDesktopFile() {
    if (!desktopApi) {
      return;
    }

    try {
      const result = await desktopApi.openMarkdownFile();
      if (result.canceled) {
        setStatusMessage('已取消打开');
        return;
      }

      const nextPath = result.filePath ?? null;
      const nextContent = result.content ?? '';
      const nextName = formatFileName(nextPath);

      resetUndoHistory();
      setDoc({
        filePath: nextPath,
        displayName: nextName,
        content: nextContent
      });
      setLastSavedContent(nextContent);
      setViewMode('preview');
      setBeautifyPreview(null);
      setTranslationPreview(null);
      setDocxImportPreview(null);
      rememberRecentFile(nextPath, nextName);
      setStatusMessage(`已打开 ${nextName}`);
      focusEditorSoon('preview');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知读取错误';
      setStatusMessage(`打开失败：${message}`);
      console.error('[renderer] open failed', error);
    }
  }

  async function handleOpenFile() {
    if (!confirmDiscardChanges()) {
      return;
    }

    if (desktopApi) {
      await openDesktopFile();
      return;
    }

    setStatusMessage('请选择要打开的 Markdown 文件');
    browserFileInputRef.current?.click();
  }

  async function handleImportDocx() {
    if (!desktopApi) {
      setStatusMessage('Word 导入目前需要在桌面应用中使用');
      return;
    }

    try {
      setIsDocxImporting(true);
      setStatusMessage('正在解析 Word 文档...');
      const result = await desktopApi.openDocxFile();

      if (result.canceled) {
        setStatusMessage('已取消导入 Word 文档');
        return;
      }

      setDocxImportPreview({
        sourceFilePath: result.sourceFilePath,
        sourceFileName: result.sourceFileName,
        markdown: result.markdown,
        previewHtml: result.previewHtml,
        messages: result.messages,
        summary: result.summary,
        imageAssets: result.imageAssets
      });
      setStatusMessage(`已生成 ${result.sourceFileName} 的转换预览`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知导入错误';
      setStatusMessage(`导入 Word 文档失败：${message}`);
      console.error('[renderer] import docx failed', error);
    } finally {
      setIsDocxImporting(false);
    }
  }

  async function importDocxAsNewDocument() {
    if (!desktopApi || !docxImportPreview) {
      return;
    }

    if (!confirmDiscardChanges()) {
      return;
    }

    try {
      setStatusMessage('正在导入为新文档...');
      const defaultPath = docxImportPreview.sourceFileName.replace(/\.(docx|docm)$/i, '.md');
      const result = await desktopApi.saveImportedMarkdownFile({
        defaultPath,
        markdown: docxImportPreview.markdown,
        imageAssets: docxImportPreview.imageAssets
      });

      if (result.canceled) {
        setStatusMessage('已取消导入为新文档');
        return;
      }

      const nextPath = result.filePath;
      const nextName = formatFileName(nextPath);
      resetUndoHistory();
      setDoc({
        filePath: nextPath,
        displayName: nextName,
        content: result.content
      });
      setLastSavedContent(result.content);
      setViewMode('preview');
      setBeautifyPreview(null);
      setTranslationPreview(null);
      setDocxImportPreview(null);
      rememberRecentFile(nextPath, nextName);
      setStatusMessage(`已导入 ${nextName}`);
      focusEditorSoon('preview');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知导入错误';
      setStatusMessage(`导入为新文档失败：${message}`);
      console.error('[renderer] import docx as new failed', error);
    }
  }

  async function insertDocxIntoCurrentDocument() {
    if (!docxImportPreview) {
      return;
    }

    try {
      let nextMarkdown = docxImportPreview.markdown;

      if (docxImportPreview.imageAssets.length > 0) {
        if (!desktopApi) {
          setStatusMessage('含图片的 Word 导入需要在桌面应用中使用');
          return;
        }

        if (!doc.filePath) {
          setStatusMessage('插入含图片的 Word 内容前，请先保存当前 Markdown 文档');
          return;
        }

        setStatusMessage('正在整理导入图片资源...');
        const materialized = await desktopApi.materializeImportedMarkdown({
          documentPath: doc.filePath,
          markdown: nextMarkdown,
          imageAssets: docxImportPreview.imageAssets
        });
        nextMarkdown = materialized.content;
      }

      const normalized = normalizeLineEndings(nextMarkdown).trim();
      const insertion = normalized ? `${normalized}\n` : '';
      const inserted = insertIntoCurrentPosition(insertion, { preferProgrammatic: true });
      if (!inserted) {
        return;
      }
      setDocxImportPreview(null);
      setStatusMessage(`已插入 ${docxImportPreview.sourceFileName} 的转换内容`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知插入错误';
      setStatusMessage(`插入 Word 转换内容失败：${message}`);
      console.error('[renderer] insert docx import failed', error);
    }
  }

  async function handleBrowserFileSelected(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = '';

    if (!file) {
      setStatusMessage('已取消打开');
      return;
    }

    try {
      const content = await file.text();
      const filePath = (file as File & { path?: string }).path ?? null;

      resetUndoHistory();
      setDoc({
        filePath,
        displayName: file.name,
        content
      });
      setLastSavedContent(content);
      setViewMode('preview');
      setBeautifyPreview(null);
      setTranslationPreview(null);
      setDocxImportPreview(null);
      rememberRecentFile(filePath, file.name);
      setStatusMessage(`已打开 ${file.name}`);
      focusEditorSoon('preview');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知读取错误';
      setStatusMessage(`打开失败：${message}`);
      console.error('[renderer] open failed', error);
    }
  }

  async function handleBrowserImageSelected(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = '';

    if (!file) {
      setStatusMessage('已取消引用图片');
      return;
    }

    try {
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ''));
        reader.onerror = () => reject(reader.error || new Error('读取图片失败'));
        reader.readAsDataURL(file);
      });

      const markdown = getImageMarkdown(dataUrl, getDisplayNameWithoutExtension(file.name));
      insertIntoCurrentPosition(markdown);
      setStatusMessage(`已引用图片 ${file.name}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知图片引用错误';
      setStatusMessage(`引用图片失败：${message}`);
    }
  }

  async function saveInBrowser() {
    const fileNameForDownload = fileName.endsWith('.md') ? fileName : `${fileName}.md`;
    const blob = new Blob([doc.content], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');

    anchor.href = url;
    anchor.download = fileNameForDownload;
    anchor.click();
    URL.revokeObjectURL(url);

    setLastSavedContent(doc.content);
    setStatusMessage(`已导出 ${fileNameForDownload}`);
  }

  async function handleSaveFile(forceDialog: boolean) {
    if (!desktopApi) {
      await saveInBrowser();
      return;
    }

    try {
      setStatusMessage(forceDialog ? '正在另存为...' : '正在保存...');
      const result = await desktopApi.saveMarkdownFile({
        filePath: doc.filePath,
        content: doc.content,
        forceDialog
      });

      if (result.canceled) {
        setStatusMessage('已取消保存');
        return;
      }

      const savedPath = result.filePath ?? doc.filePath;
      const savedName = formatFileName(savedPath ?? null);
      setDoc((current) => ({
        ...current,
        filePath: savedPath ?? null,
        displayName: savedName
      }));
      setLastSavedContent(doc.content);
      rememberRecentFile(savedPath ?? null, savedName);
      setStatusMessage(`已保存 ${savedName}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知保存错误';
      setStatusMessage(`保存失败：${message}`);
      console.error('[renderer] save failed', error);
    }
  }

  async function handleExportHtml() {
    const defaultName = fileName.replace(/\.(md|markdown|mdown|txt)$/i, '') || '未命名';
    const exportBodyHtml = sanitizeHtml(renderMarkdown(doc.content, false));
    const html = buildExportHtml(defaultName, exportBodyHtml);

    if (!desktopApi) {
      const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `${defaultName}.html`;
      anchor.click();
      URL.revokeObjectURL(url);
      setStatusMessage(`已导出 ${defaultName}.html`);
      return;
    }

    try {
      const result = await desktopApi.exportHtmlFile({
        defaultPath: `${defaultName}.html`,
        html
      });

      if (result.canceled) {
        setStatusMessage('已取消导出');
        return;
      }

      setStatusMessage(`已导出 ${formatFileName(result.filePath ?? null)}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知导出错误';
      setStatusMessage(`导出失败：${message}`);
    }
  }

  async function handleExportPdf() {
    const defaultName = `${fileName.replace(/\.(md|markdown|mdown|txt)$/i, '') || '未命名'}.pdf`;

    if (!desktopApi) {
      setStatusMessage('PDF 导出需要在桌面应用中使用');
      return;
    }

    try {
      const result = await desktopApi.exportPdfFile({
        defaultPath: defaultName
      });

      if (result.canceled) {
        setStatusMessage('已取消导出');
        return;
      }

      setStatusMessage(`已导出 ${formatFileName(result.filePath ?? null)}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知导出错误';
      setStatusMessage(`导出 PDF 失败：${message}`);
    }
  }

  async function handleReferenceImage() {
    if (!desktopApi) {
      browserImageInputRef.current?.click();
      return;
    }

    try {
      const result = await desktopApi.pickImageReference({
        documentPath: doc.filePath
      });

      if (result.canceled) {
        setStatusMessage('已取消引用图片');
        return;
      }

      const markdown = getImageMarkdown(
        result.markdownPath,
        getDisplayNameWithoutExtension(result.displayName)
      );
      insertIntoCurrentPosition(markdown);
      setStatusMessage(`已引用图片 ${result.displayName}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知图片引用错误';
      setStatusMessage(`引用图片失败：${message}`);
    }
  }

  function openBeautifyPreview(strength = beautifyStrength) {
    const selection = getPreferredSelection();
    const useSelection = Boolean(selection && selection.text.trim());
    const sourceText = useSelection ? selection!.text : doc.content;
    const result = beautifyDocumentContent(sourceText, strength);
    const original = normalizeLineEndings(sourceText);
    const beautified = normalizeLineEndings(result.content);

    setBeautifyPreview({
      scope: useSelection ? 'selection' : 'document',
      strength,
      original,
      beautified,
      summary: result.summary,
      changes: result.changes,
      hasChanges: original !== beautified,
      selectionRange: useSelection && selection ? { start: selection.start, end: selection.end } : null
    });

    setStatusMessage(useSelection ? '已生成选中内容的 AI 美化预览' : '已生成整篇文档的 AI 美化预览');
  }

  function applyBeautifyPreview() {
    if (!beautifyPreview) {
      return;
    }

    recordProgrammaticDocumentChange();

    if (beautifyPreview.scope === 'selection' && beautifyPreview.selectionRange) {
      setDoc((current) => ({
        ...current,
        content: `${current.content.slice(0, beautifyPreview.selectionRange.start)}${beautifyPreview.beautified}${current.content.slice(beautifyPreview.selectionRange.end)}`
      }));
      lastSelectionRef.current = {
        start: beautifyPreview.selectionRange.start,
        end: beautifyPreview.selectionRange.start + beautifyPreview.beautified.length,
        source: 'editor'
      };
      focusEditorSoon('preview', beautifyPreview.selectionRange.start + beautifyPreview.beautified.length);
      setStatusMessage(`AI 已优化当前选区：${beautifyPreview.summary}`);
    } else {
      setDoc((current) => ({
        ...current,
        content: beautifyPreview.beautified
      }));
      focusEditorSoon('preview');
      setStatusMessage(`AI 已优化整篇文档：${beautifyPreview.summary}`);
    }

    setBeautifyPreview(null);
  }

  function handleBeautifyStrengthChange(strength: BeautifyStrength) {
    setBeautifyStrength(strength);
    openBeautifyPreview(strength);
  }

  function handleContextMenu(event: ReactMouseEvent<HTMLElement>) {
    event.preventDefault();

    try {
      const root = previewEditorRef.current;
      if (root) {
        const pointRange = resolveRangeFromPoint(event.clientX, event.clientY);
        if (pointRange && root.contains(pointRange.startContainer)) {
          applyDomRangeSelection(pointRange);
          updateEditorSelectionState('preview');
        } else if (root.contains(event.target as Node)) {
          moveCaretToNearestLineEnd(root, event.target, event.clientX, event.clientY);
          updateEditorSelectionState('preview');
        }

        const sectionFromPoint = getSectionRangeFromClientPoint(event.clientX, event.clientY);
        if (sectionFromPoint?.id) {
          lastActiveSectionIdRef.current = sectionFromPoint.id;
          setLiveAiSectionId(sectionFromPoint.id);
        }
      }
    } catch (error) {
      console.error('[renderer] context menu sync failed', error);
    }

    const viewportPadding = 6;
    const estimatedWidth = Math.min(360, Math.max(220, window.innerWidth - viewportPadding * 2));
    const estimatedHeight = Math.min(520, Math.max(220, window.innerHeight - viewportPadding * 2));
    const nextX = Math.max(viewportPadding, Math.min(event.clientX, window.innerWidth - estimatedWidth - viewportPadding));
    const nextY = Math.max(viewportPadding, Math.min(event.clientY, window.innerHeight - estimatedHeight - viewportPadding));

    setContextMenu({
      x: nextX,
      y: nextY
    });
  }

  function runContextAction(action: () => void | Promise<void>) {
    setContextMenu(null);
    void action();
  }

  function applyTheme(nextTheme: ThemeName) {
    setTheme(nextTheme);
    const labels = {
      paper: '米白主题',
      mist: '雾蓝主题',
      slate: '灰砚主题',
      graphite: '石墨主题',
      terminal: '终端绿主题',
      nightcode: '夜码蓝主题',
      campus: '课堂蓝主题',
      youth: '青春橙主题'
    } satisfies Record<ThemeName, string>;
    setStatusMessage(`已切换到${labels[nextTheme]}`);
  }

  function updateAiProfile(profileId: string, updater: (profile: AIProfile) => AIProfile) {
    setAiProfiles((current) => current.map((profile) => (profile.id === profileId ? updater(profile) : profile)));
  }

  function updateActiveAiProfile(updater: (profile: AIProfile) => AIProfile) {
    updateAiProfile(activeAiProfile.id, updater);
  }

  function createProfileName(baseLabel: string) {
    const usedNames = new Set(aiProfiles.map((profile) => profile.name));
    if (!usedNames.has(baseLabel)) {
      return baseLabel;
    }

    let index = 2;
    while (usedNames.has(`${baseLabel} ${index}`)) {
      index += 1;
    }

    return `${baseLabel} ${index}`;
  }

  function addAiProfile(presetKey: AIPresetKey = 'custom') {
    const preset = aiProviderPresets[presetKey];
    const nextProfile = createAiProfile(createProfileName(preset.label), {
      provider: preset.provider,
      apiBase: preset.apiBase,
      apiKey: '',
      model: preset.model,
      temperature: aiSettings.temperature,
      maxTokens: aiSettings.maxTokens
    });

    setAiProfiles((current) => [...current, nextProfile]);
    setActiveAiProfileId(nextProfile.id);
    setIsAiConfigOpen(true);
    setAiConnectionMessage(
      presetKey === 'builtin'
        ? '已新增内置文档助手配置。'
        : `已新增 ${preset.label} 配置，请填写 API Key 后测试连接。`
    );
    setStatusMessage(`已新增 ${preset.label} 配置`);
  }

  function duplicateActiveAiProfile() {
    const duplicated = createAiProfile(createProfileName(`${activeAiProfile.name} 副本`), {
      provider: activeAiProfile.provider,
      apiBase: activeAiProfile.apiBase,
      apiKey: activeAiProfile.apiKey,
      model: activeAiProfile.model,
      temperature: activeAiProfile.temperature,
      maxTokens: activeAiProfile.maxTokens
    });

    setAiProfiles((current) => [...current, duplicated]);
    setActiveAiProfileId(duplicated.id);
    setIsAiConfigOpen(true);
    setStatusMessage(`已复制配置 ${activeAiProfile.name}`);
  }

  function deleteActiveAiProfile() {
    if (aiProfiles.length <= 1) {
      setStatusMessage('至少保留一个 AI 配置');
      return;
    }

    const currentIndex = aiProfiles.findIndex((profile) => profile.id === activeAiProfile.id);
    const fallback = aiProfiles[currentIndex - 1] ?? aiProfiles[currentIndex + 1] ?? aiProfiles[0];
    setAiProfiles((current) => current.filter((profile) => profile.id !== activeAiProfile.id));
    if (fallback) {
      setActiveAiProfileId(fallback.id);
    }
    setStatusMessage(`已删除配置 ${activeAiProfile.name}`);
  }

  function applyAiPreset(presetKey: AIPresetKey) {
    const preset = aiProviderPresets[presetKey];

    updateActiveAiProfile((current) => ({
      ...current,
      provider: preset.provider,
      apiBase: preset.apiBase,
      model: preset.model
    }));

    if (presetKey === 'builtin') {
      setAiConnectionMessage('当前使用内置文档助手，无需远程连接。');
      setStatusMessage('已切换到内置文档助手');
      return;
    }

    if (presetKey === 'custom') {
      setAiConnectionMessage('已切换到自定义兼容接口，请填写 Base URL、API Key 和模型名称。');
      setStatusMessage('已切换到自定义兼容接口');
      return;
    }

    setAiConnectionMessage(`已载入 ${preset.label} 预设，请补充 API Key 后测试连接。`);
    setStatusMessage(`已载入 ${preset.label} 模型预设`);
  }

  function restoreAiPresetDefaults() {
    if (currentAiPreset === 'builtin' || currentAiPreset === 'custom') {
      setStatusMessage(currentAiPreset === 'builtin' ? '内置文档助手无需恢复默认值' : '当前是自定义接口，没有可恢复的推荐模型');
      return;
    }

    const preset = aiProviderPresets[currentAiPreset];
    updateActiveAiProfile((current) => ({
      ...current,
      provider: preset.provider,
      apiBase: preset.apiBase,
      model: preset.model
    }));
    setAiConnectionMessage(`已恢复 ${preset.label} 的推荐接口和模型，请补充 API Key 后测试连接。`);
    setStatusMessage(`已恢复 ${preset.label} 默认模型`);
  }

  async function openPresetConsole() {
    const preset = aiProviderPresets[currentAiPreset];
    if (!preset.consoleUrl) {
      setStatusMessage('当前预设没有可打开的控制台链接');
      return;
    }

    try {
      if (desktopApi?.openExternalLink) {
        await desktopApi.openExternalLink(preset.consoleUrl);
      } else {
        window.open(preset.consoleUrl, '_blank', 'noopener,noreferrer');
      }
      setStatusMessage(`已打开 ${preset.label} 控制台`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知打开错误';
      setStatusMessage(`打开链接失败：${message}`);
    }
  }

  function getAiProviderLabel() {
    const canUseRemote =
      aiSettings.provider === 'openai-compatible' &&
      aiSettings.apiBase.trim() &&
      aiSettings.apiKey.trim() &&
      aiSettings.model.trim();

    return canUseRemote ? '自定义兼容 AI 模型' : '内置文档助手';
  }

  const currentAiPreset = inferAiPreset(aiSettings);
  const currentAiPresetConfig = aiProviderPresets[currentAiPreset];

  function appendAiRequestLog(entry: Omit<AIRequestLog, 'id' | 'time'>) {
    const logEntry: AIRequestLog = {
      ...entry,
      id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
      time: new Date().toLocaleTimeString('zh-CN', { hour12: false })
    };

    const prefix = `[ai:${entry.kind}:${entry.status}]`;
    if (entry.status === 'error') {
      console.error(prefix, entry.detail, entry.endpoint);
    } else {
      console.log(prefix, entry.detail, entry.endpoint);
    }

    setAiRequestLogs((current) => [logEntry, ...current].slice(0, 24));
  }

  function validateRemoteAiSettings() {
    if (aiSettings.provider !== 'openai-compatible') {
      return {
        ok: false as const,
        message: '当前提供方不是 OpenAI-compatible，正在使用内置文档助手。'
      };
    }

    if (!aiSettings.apiBase.trim()) {
      return {
        ok: false as const,
        message: 'AI 配置不完整：缺少 API Base URL。'
      };
    }

    if (!aiSettings.apiKey.trim()) {
      return {
        ok: false as const,
        message: 'AI 配置不完整：缺少 API Key。'
      };
    }

    if (!aiSettings.model.trim()) {
      return {
        ok: false as const,
        message: 'AI 配置不完整：缺少模型名称。'
      };
    }

    return {
      ok: true as const,
      endpoint: getOpenAiCompatibleEndpoint(aiSettings.apiBase)
    };
  }

  async function handleTestAiConnection() {
    const validation = validateRemoteAiSettings();

    if (!validation.ok) {
      setAiToastMessage(null);
      setAiConnectionMessage(validation.message);
      appendAiRequestLog({
        kind: 'connection',
        status: 'skipped',
        providerLabel: getAiProviderLabel(),
        endpoint: validation.ok ? validation.endpoint : 'builtin://local-assistant',
        detail: validation.message
      });
      setStatusMessage(validation.message);
      return;
    }

    setIsAiTesting(true);
    setAiToastMessage(null);
    setAiConnectionMessage('正在测试远程 AI 连接...');
    appendAiRequestLog({
      kind: 'connection',
      status: 'started',
      providerLabel: '自定义兼容 AI 模型',
      endpoint: validation.endpoint,
      detail: `开始测试连接，模型：${aiSettings.model}`
    });

    try {
      const result = await testOpenAiCompatibleConnection(aiSettings);
      const detail = `连接成功，模型返回：${result}`;
      setAiConnectionMessage(detail);
      setAiToastMessage('AI 连接测试成功');
      appendAiRequestLog({
        kind: 'connection',
        status: 'success',
        providerLabel: '自定义兼容 AI 模型',
        endpoint: validation.endpoint,
        detail
      });
      setStatusMessage('AI 连接测试成功');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知连接错误';
      setAiToastMessage(null);
      setAiConnectionMessage(`连接失败：${message}`);
      appendAiRequestLog({
        kind: 'connection',
        status: 'error',
        providerLabel: '自定义兼容 AI 模型',
        endpoint: validation.endpoint,
        detail: `连接失败：${message}`
      });
      setStatusMessage(`AI 连接失败：${message}`);
    } finally {
      setIsAiTesting(false);
    }
  }

  async function generateAiActionResult(
    action: AIActionKind,
    sourceText: string,
    scope: AIContextScope
  ) {
    const validation = validateRemoteAiSettings();

    if (!validation.ok) {
      appendAiRequestLog({
        kind: 'action',
        status: 'skipped',
        providerLabel: '内置文档助手',
        endpoint: 'builtin://local-assistant',
        detail: validation.message
      });
      return {
        content: transformWithBuiltinAi(action, sourceText),
        providerLabel: '内置文档助手'
      };
    }

    const actionLabelMap: Record<AIActionKind, string> = {
      rewrite: '改写',
      expand: '扩写',
      shorten: '精简',
      continue: '续写'
    };

    const systemPrompt =
      '你是一个帮助用户编辑 Markdown 技术文档的 AI 助手。请严格保留 Markdown 结构，返回可直接写回文档的 Markdown 内容，不要输出额外解释。';
    const userPrompt = [
      `当前任务：${actionLabelMap[action]}`,
      `上下文范围：${scope === 'selection' ? '当前选区' : scope === 'section' ? '当前章节' : '整篇文档'}`,
      '',
      '请处理以下 Markdown 内容：',
      sourceText
    ].join('\n');

    appendAiRequestLog({
      kind: 'action',
      status: 'started',
      providerLabel: '自定义兼容 AI 模型',
      endpoint: validation.endpoint,
      detail: `开始执行 AI ${actionLabelMap[action]}，上下文：${scope}`
    });

    const content = await requestOpenAiCompatibleCompletion(aiSettings, systemPrompt, userPrompt);
    appendAiRequestLog({
      kind: 'action',
      status: 'success',
      providerLabel: '自定义兼容 AI 模型',
      endpoint: validation.endpoint,
      detail: `AI ${actionLabelMap[action]}完成，返回 ${content.length} 个字符`
    });

    return {
      content,
      providerLabel: '自定义兼容 AI 模型'
    };
  }

  async function handleAiAction(action: AIActionKind) {
    const context = resolveAiContext(aiContextScope);
    if (!context.text.trim()) {
      setStatusMessage('当前没有可供 AI 处理的内容');
      return;
    }

    setIsAiLoading(true);
    try {
      const result = await generateAiActionResult(action, context.text, context.scope);
      const nextContent = normalizeLineEndings(result.content);
      setAiActionPreview({
        action,
        scope: context.scope,
        original: normalizeLineEndings(context.text),
        result: nextContent,
        providerLabel: result.providerLabel,
        hasChanges: normalizeLineEndings(context.text) !== nextContent,
        selectionRange: context.selectionRange,
        sectionId: context.sectionId,
        sectionRange: context.sectionRange
      });
      setStatusMessage(`已生成 ${result.providerLabel} 的${action === 'rewrite' ? '改写' : action === 'expand' ? '扩写' : action === 'shorten' ? '精简' : '续写'}预览`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知 AI 错误';
      setStatusMessage(`AI 处理失败：${message}`);
    } finally {
      setIsAiLoading(false);
    }
  }

  async function handleSendAiMessage() {
    const prompt = aiInput.trim();
    if (!prompt || isAiLoading) {
      return;
    }

    const context = resolveAiContext(aiContextScope);
    const userMessage: AIChatMessage = {
      id: `user-${Date.now()}`,
      role: 'user',
      content: prompt,
      contextScope: context.scope,
      selectionRange: context.selectionRange,
      sectionId: context.sectionId,
      sectionRange: context.sectionRange
    };

    setAiMessages((current) => [...current, userMessage]);
    setAiInput('');
    setIsAiLoading(true);

    try {
      const validation = validateRemoteAiSettings();
      let providerLabel = '内置文档助手';
      let reply = '';

      if (!validation.ok) {
        appendAiRequestLog({
          kind: 'chat',
          status: 'skipped',
          providerLabel,
          endpoint: 'builtin://local-assistant',
          detail: validation.message
        });
        reply = buildBuiltinAssistantReply(prompt, context.text, context.scope, fileName);
      } else {
        providerLabel = '自定义兼容 AI 模型';
        appendAiRequestLog({
          kind: 'chat',
          status: 'started',
          providerLabel,
          endpoint: validation.endpoint,
          detail: `开始对话，请求上下文：${context.scope}`
        });
        reply = await requestOpenAiCompatibleCompletion(
          aiSettings,
          '你是一个帮助用户编写 Markdown 技术文档的 AI 助手。请优先围绕给定文档上下文回答，输出简洁、结构化、可执行的建议。不要编造没有出现在上下文中的事实。',
          [
            `文件名：${fileName}`,
            `上下文范围：${context.scope === 'selection' ? '当前选区' : context.scope === 'section' ? '当前章节' : '整篇文档'}`,
            context.sectionId
              ? `当前父标题级别：${'#'.repeat(getOutlineItemById(context.sectionId)?.level ?? 1)}（请确保你新生成的标题级别低于父标题，不要使用同级或更高等级标题）`
              : '如果需要生成标题，请保持标题层级与上下文结构一致，不要越级。',
            '',
            '文档上下文：',
            context.text,
            '',
            '用户问题：',
            prompt
          ].join('\n')
        );
        appendAiRequestLog({
          kind: 'chat',
          status: 'success',
          providerLabel,
          endpoint: validation.endpoint,
          detail: `AI 对话完成，返回 ${reply.length} 个字符`
        });
      }

      setAiMessages((current) => [
        ...current,
        {
          id: `assistant-${Date.now()}`,
          role: 'assistant',
          content: reply,
          contextScope: context.scope,
          providerLabel,
          selectionRange: context.selectionRange,
          sectionId: context.sectionId,
          sectionRange: context.sectionRange
        }
      ]);
      setStatusMessage(`AI 已完成回答（${providerLabel}）`);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知 AI 错误';
      const validation = validateRemoteAiSettings();
      appendAiRequestLog({
        kind: 'chat',
        status: 'error',
        providerLabel: validation.ok ? '自定义兼容 AI 模型' : '内置文档助手',
        endpoint: validation.ok ? validation.endpoint : 'builtin://local-assistant',
        detail: `AI 对话失败：${message}`
      });
      setStatusMessage(`AI 对话失败：${message}`);
    } finally {
      setIsAiLoading(false);
    }
  }

  async function copyText(text: string, successMessage: string) {
    try {
      await navigator.clipboard.writeText(text);
      setStatusMessage(successMessage);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知复制错误';
      setStatusMessage(`复制失败：${message}`);
    }
  }

  function appendToCurrentSection(
    text: string,
    targetSection?: { start: number; end: number } | null,
    targetSectionId?: string | null
  ) {
    const section = getSectionRangeByOutlineId(targetSectionId ?? null) ?? targetSection ?? getPreferredAiSectionRange();
    if (!section) {
      insertIntoCurrentPosition(text, { preferProgrammatic: true, sectionId: targetSectionId ?? effectiveAiSectionId });
      return;
    }

    const normalized = prepareAiMarkdownForSection(text, targetSectionId ?? section.id).trim();
    const prefix = doc.content.slice(0, section.end).endsWith('\n') ? '\n' : '\n\n';
    const nextText = `${prefix}${normalized}\n`;
    replaceContentRange(section.end, section.end, nextText, 'preview', {
      start: section.end + nextText.length,
      end: section.end + nextText.length
    });
  }

  function applyAiActionPreview(mode: 'insert' | 'replace' | 'append' | 'copy') {
    if (!aiActionPreview) {
      return;
    }

    if (mode === 'copy') {
      void copyText(aiActionPreview.result, '已复制 AI 结果');
      return;
    }

    if (mode === 'insert') {
      insertIntoCurrentPosition(aiActionPreview.result, {
        preferProgrammatic: true,
        sectionId: aiActionPreview.sectionId ?? effectiveAiSectionId
      });
      setAiActionPreview(null);
      setStatusMessage('已插入 AI 结果');
      return;
    }

    if (mode === 'append') {
      appendToCurrentSection(aiActionPreview.result, aiActionPreview.sectionRange, aiActionPreview.sectionId);
      setAiActionPreview(null);
      setStatusMessage('已追加到当前章节');
      return;
    }

    if (aiActionPreview.selectionRange) {
      replaceContentRange(
        aiActionPreview.selectionRange.start,
        aiActionPreview.selectionRange.end,
        aiActionPreview.result,
        'preview',
        {
          start: aiActionPreview.selectionRange.start,
          end: aiActionPreview.selectionRange.start + aiActionPreview.result.length
        }
      );
      setAiActionPreview(null);
      setStatusMessage('已替换当前选区');
      return;
    }

    setStatusMessage('当前没有可替换的选区，已保留预览结果');
  }

  async function openTranslationPreview(direction = translationDirection) {
    const selection = getPreferredSelection();
    const useSelection = Boolean(selection && selection.text.trim());
    const sourceText = useSelection ? selection!.text : doc.content;

    setTranslationDirection(direction);
    setIsTranslating(true);
    try {
      const validation = validateRemoteAiSettings();
      if (validation.ok) {
        appendAiRequestLog({
          kind: 'translate',
          status: 'started',
          providerLabel: '自定义兼容 AI 模型',
          endpoint: validation.endpoint,
          detail: `开始翻译，方向：${direction}`
        });
      } else {
        appendAiRequestLog({
          kind: 'translate',
          status: 'skipped',
          providerLabel: '内置文档助手',
          endpoint: 'builtin://local-assistant',
          detail: validation.message
        });
      }

      const result = await translateDocumentContent(sourceText, direction, aiSettings);
      const translated = normalizeLineEndings(result.translated);
      setTranslationPreview({
        scope: useSelection ? 'selection' : 'document',
        direction,
        original: normalizeLineEndings(sourceText),
        translated,
        providerLabel: result.providerLabel,
        hasChanges: normalizeLineEndings(sourceText) !== translated,
        selectionRange: useSelection && selection ? { start: selection.start, end: selection.end } : null
      });
      appendAiRequestLog({
        kind: 'translate',
        status: 'success',
        providerLabel: result.providerLabel,
        endpoint: validation.ok ? validation.endpoint : 'builtin://local-assistant',
        detail: `翻译完成，返回 ${translated.length} 个字符`
      });
      setStatusMessage(useSelection ? '已生成选中内容的翻译预览' : '已生成整篇文档的翻译预览');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知翻译错误';
      const validation = validateRemoteAiSettings();
      appendAiRequestLog({
        kind: 'translate',
        status: 'error',
        providerLabel: validation.ok ? '自定义兼容 AI 模型' : '内置文档助手',
        endpoint: validation.ok ? validation.endpoint : 'builtin://local-assistant',
        detail: `翻译失败：${message}`
      });
      setStatusMessage(`翻译失败：${message}`);
    } finally {
      setIsTranslating(false);
    }
  }

  function applyTranslationPreview() {
    if (!translationPreview) {
      return;
    }

    recordProgrammaticDocumentChange();

    if (translationPreview.scope === 'selection' && translationPreview.selectionRange) {
      setDoc((current) => ({
        ...current,
        content: `${current.content.slice(0, translationPreview.selectionRange.start)}${translationPreview.translated}${current.content.slice(translationPreview.selectionRange.end)}`
      }));
      focusEditorSoon('preview', translationPreview.selectionRange.start + translationPreview.translated.length);
    } else {
      setDoc((current) => ({
        ...current,
        content: translationPreview.translated
      }));
      focusEditorSoon('preview');
    }

    setTranslationPreview(null);
    setStatusMessage(`已应用翻译结果（${translationPreview.providerLabel}）`);
  }

  async function handleOpenRecentFile(file: RecentFile) {
    if (!desktopApi) {
      setStatusMessage('最近文件功能需要在桌面应用中使用');
      return;
    }

    if (!confirmDiscardChanges()) {
      return;
    }

    try {
      const result = await desktopApi.readMarkdownFile(file.filePath);
      resetUndoHistory();
      setDoc({
        filePath: result.filePath,
        displayName: formatFileName(result.filePath, file.displayName),
        content: result.content
      });
      setLastSavedContent(result.content);
      setViewMode('preview');
      setBeautifyPreview(null);
      rememberRecentFile(result.filePath, file.displayName);
      setStatusMessage(`已重新打开 ${file.displayName}`);
      focusEditorSoon('preview');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知读取错误';
      setStatusMessage(`打开最近文件失败：${message}`);
      setRecentFiles((current) => current.filter((item) => item.filePath !== file.filePath));
    }
  }

  function jumpToHeading(id: string) {
    setLiveAiSectionId(id);
    lastActiveSectionIdRef.current = id;
    window.setTimeout(() => {
      const root = previewEditorRef.current;
      const target = root?.querySelector<HTMLElement>(`[data-outline-id="${id}"]`);
      if (!root || !target) {
        return;
      }

      root.focus();
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 50);
  }

  async function insertImageFile(file: File) {
    if (!file.type.startsWith('image/')) {
      return;
    }

    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ''));
      reader.onerror = () => reject(reader.error || new Error('读取图片失败'));
      reader.readAsDataURL(file);
    });

    let markdown = '';

    if (!desktopApi) {
      markdown = getImageMarkdown(dataUrl, file.name.replace(/\.[^.]+$/, ''));
      insertIntoCurrentPosition(markdown);
      setStatusMessage(`已插入图片 ${file.name}`);
      return;
    }

    const saved = await desktopApi.saveImageAsset({
      documentPath: doc.filePath,
      originalName: file.name,
      dataUrl
    });

    if (saved.canceled || !saved.filePath) {
      setStatusMessage('已取消插入图片');
      return;
    }

    const documentDir = doc.filePath
      ? doc.filePath.slice(0, Math.max(doc.filePath.lastIndexOf('\\'), doc.filePath.lastIndexOf('/')) + 1)
      : '';
    const relativePath =
      documentDir && saved.filePath.startsWith(documentDir)
        ? saved.filePath.slice(documentDir.length).replace(/\\/g, '/')
        : saved.filePath.replace(/\\/g, '/');
    markdown = getImageMarkdown(relativePath, file.name.replace(/\.[^.]+$/, ''));

    insertIntoCurrentPosition(markdown);
    setStatusMessage(`已插入图片 ${file.name}`);
  }

  async function handlePaste(event: ClipboardEvent<HTMLDivElement>) {
    const files = Array.from(event.clipboardData.files);
    const imageFile = files.find((file) => file.type.startsWith('image/'));
    if (!imageFile) {
      return;
    }

    event.preventDefault();
    try {
      await insertImageFile(imageFile);
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知图片插入错误';
      setStatusMessage(`插入图片失败：${message}`);
    }
  }

  async function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    const imageFiles = Array.from(event.dataTransfer.files).filter((file) => file.type.startsWith('image/'));

    for (const file of imageFiles) {
      try {
        await insertImageFile(file);
      } catch (error) {
        const message = error instanceof Error ? error.message : '未知图片插入错误';
        setStatusMessage(`插入图片失败：${message}`);
        break;
      }
    }
  }

  function handleRichEditorInput(event: FormEvent<HTMLDivElement>) {
    const root = event.currentTarget;
    const markdown = serializeRichEditorContent(root);
    if (markdown === doc.content) {
      return;
    }

    pushUndoHistoryEntry({
      content: doc.content,
      selection: getContentEditableSelectionOffsets(root)
    });
    redoHistoryRef.current = [];
    lastSelectionRef.current = null;
    lastRenderedMarkdownRef.current = markdown;
    setDoc((current) => ({ ...current, content: markdown }));
  }

  function handleEditorSelectionEvent(
    source: SelectionSource,
    event:
      | ReactMouseEvent<HTMLDivElement>
      | KeyboardEvent<HTMLDivElement>
      | FormEvent<HTMLDivElement>
  ) {
    if ('clientX' in event && 'clientY' in event) {
      const sectionFromPoint = getSectionRangeFromClientPoint(event.clientX, event.clientY);
      if (sectionFromPoint?.id) {
        lastActiveSectionIdRef.current = sectionFromPoint.id;
        setLiveAiSectionId(sectionFromPoint.id);
      }

      if (event.type === 'click') {
        const pointRange = resolveRangeFromPoint(event.clientX, event.clientY);
        let finalRange: Range | null = null;

        if (pointRange && event.currentTarget.contains(pointRange.startContainer)) {
          applyDomRangeSelection(pointRange);
          finalRange = pointRange;
        } else {
          moveCaretToNearestLineEnd(event.currentTarget, event.target, event.clientX, event.clientY);
          const selection = window.getSelection();
          if (selection && selection.rangeCount > 0) {
            finalRange = selection.getRangeAt(0).cloneRange();
          }
        }

        if (finalRange && event.currentTarget.contains(finalRange.startContainer)) {
          const collapsedRange = finalRange.cloneRange();
          collapsedRange.collapse(false);
          lastCaretDomRangeRef.current = collapsedRange;
          lastSelectedDomRangeRef.current = null;
          const offsets = getContentEditableRangeOffsets(event.currentTarget, collapsedRange);
          lastCaretRangeRef.current = offsets ? { ...offsets, source } : null;
          lastSelectionRef.current = null;
        }

        return;
      } else {
        moveCaretToNearestLineEnd(event.currentTarget, event.target, event.clientX, event.clientY);
      }
    }
    updateEditorSelectionState(source);
  }

  const contextMenuGroups: ContextMenuGroup[] = [
    {
      key: 'ctx-headings',
      title: '标题',
      columns: 3,
      items: [
        { key: 'ctx-h1', label: 'H1', onClick: () => insertHeadingAtCurrentPosition(1) },
        { key: 'ctx-h2', label: 'H2', onClick: () => insertHeadingAtCurrentPosition(2) },
        { key: 'ctx-h3', label: 'H3', onClick: () => insertHeadingAtCurrentPosition(3) },
        { key: 'ctx-h4', label: 'H4', onClick: () => insertHeadingAtCurrentPosition(4) },
        { key: 'ctx-h5', label: 'H5', onClick: () => insertHeadingAtCurrentPosition(5) },
        { key: 'ctx-h6', label: 'H6', onClick: () => insertHeadingAtCurrentPosition(6) }
      ]
    },
    {
      key: 'ctx-inline',
      title: '行内语法',
      columns: 2,
      items: [
        { key: 'ctx-bold', label: '加粗', onClick: () => insertInlineMarkdown('**', '**', '加粗文本', '已插入加粗语法') },
        { key: 'ctx-italic', label: '斜体', onClick: () => insertInlineMarkdown('*', '*', '斜体文本', '已插入斜体语法') },
        { key: 'ctx-bold-italic', label: '粗斜体', onClick: () => insertInlineMarkdown('***', '***', '粗斜体文本', '已插入粗斜体语法') },
        { key: 'ctx-strike', label: '删除线', onClick: () => insertInlineMarkdown('~~', '~~', '删除线文本', '已插入删除线语法') },
        { key: 'ctx-inline-code', label: '行内代码', onClick: () => insertInlineMarkdown('`', '`', 'code', '已插入行内代码') },
        { key: 'ctx-link', label: '链接', onClick: () => insertLinkAtCurrentPosition() }
      ]
    },
    {
      key: 'ctx-block',
      title: '块级语法',
      columns: 2,
      items: [
        { key: 'ctx-quote', label: '引用块', onClick: () => insertPrefixedLinesAtCurrentPosition(['引用内容'], (line) => `> ${line || '引用内容'}`, '已插入引用块') },
        { key: 'ctx-ul', label: '无序列表', onClick: () => insertPrefixedLinesAtCurrentPosition(['列表项 1', '列表项 2'], (line) => `- ${line || '列表项'}`, '已插入无序列表') },
        { key: 'ctx-ol', label: '有序列表', onClick: () => insertPrefixedLinesAtCurrentPosition(['列表项 1', '列表项 2'], (line, index) => `${index + 1}. ${line || `列表项 ${index + 1}`}`, '已插入有序列表') },
        { key: 'ctx-task', label: '任务列表', onClick: () => insertPrefixedLinesAtCurrentPosition(['待办事项 1', '待办事项 2'], (line) => `- [ ] ${line || '待办事项'}`, '已插入任务列表') },
        { key: 'ctx-task-done', label: '已完成任务', onClick: () => insertPrefixedLinesAtCurrentPosition(['已完成事项'], (line) => `- [x] ${line || '已完成事项'}`, '已插入已完成任务') },
        { key: 'ctx-code', label: '代码块', onClick: () => insertCodeBlockAtCurrentPosition() },
        { key: 'ctx-table', label: '表格', onClick: () => insertTableAtCurrentPosition() },
        { key: 'ctx-divider', label: '分割线', onClick: () => insertDividerAtCurrentPosition() }
      ]
    },
    {
      key: 'ctx-media-tools',
      title: '媒体与工具',
      columns: 2,
      items: [
        { key: 'ctx-image-syntax', label: '图片语法', onClick: () => insertImageSyntaxAtCurrentPosition() },
        { key: 'ctx-ref-image', label: '引用图片', onClick: () => handleReferenceImage() },
        { key: 'ctx-translate', label: '翻译', onClick: () => openTranslationPreview() },
        { key: 'ctx-divider-secondary', label: '分割线', onClick: () => insertDividerAtCurrentPosition() }
      ]
    }
  ];

  return (
    <div className="shell" data-theme={theme}>
      {aiToastMessage && <div className="ai-toast">{aiToastMessage}</div>}
      <input
        ref={browserFileInputRef}
        type="file"
        accept=".md,.markdown,.mdown,.txt,text/plain"
        className="visually-hidden"
        onChange={(event) => void handleBrowserFileSelected(event)}
      />
      <input
        ref={browserImageInputRef}
        type="file"
        accept="image/*"
        className="visually-hidden"
        onChange={(event) => void handleBrowserImageSelected(event)}
      />

      <aside className="sidebar sidebar--outline">
        <div className="brand-block brand-block--compact">
          <span className="eyebrow">墨笺</span>
          <h1>墨笺</h1>
        </div>

        <div className="sidebar-header">
          <h2>大纲</h2>
        </div>

        <div className="sidebar-card sidebar-card--outline">
          <div className="sidebar-card-heading">
            <span className="card-label">标题目录</span>
            <span className="sidebar-count">{outline.length} 项</span>
          </div>
          {outline.length === 0 && <span>添加 Markdown 标题后，这里会自动生成目录。</span>}
          {outline.map((item) => (
            <button
              key={item.id}
              type="button"
              className={`sidebar-link sidebar-link--outline level-${item.level} ${shouldHighlightAiSection && effectiveAiSectionId === item.id ? 'is-active' : ''}`}
              onClick={() => jumpToHeading(item.id)}
            >
              <strong>{item.text}</strong>
            </button>
          ))}
        </div>
      </aside>

      <main className="workspace">
        <div className={`workspace-body ${isAiDrawerOpen ? 'is-ai-open' : ''}`}>
          <section className="editor-grid editor-grid--preview" onContextMenu={handleContextMenu}>
            <article className="panel panel--workspace">
              <div className="panel-heading">
                <span className="panel-title">实时预览编辑</span>
                <button type="button" className="panel-toggle" onClick={() => setIsAiDrawerOpen((current) => !current)}>
                  {isAiDrawerOpen ? '关闭 AI 会话面板' : 'AI 会话面板'}
                </button>
              </div>

              <div
                ref={previewEditorRef}
                className="markdown-body rich-editor preview-editor--workspace preview-editor--single"
                contentEditable
                suppressContentEditableWarning
                data-placeholder="在这里直接编辑富文本 Markdown 内容..."
                onInput={handleRichEditorInput}
                onClick={(event) => handleEditorSelectionEvent('preview', event)}
                onMouseUp={(event) => handleEditorSelectionEvent('preview', event)}
                onKeyDown={(event) => {
                  if ((event.ctrlKey || event.metaKey) && !event.altKey) {
                    if (handleUndoRedoShortcut(event.key, event.shiftKey)) {
                      event.preventDefault();
                      event.stopPropagation();
                      return;
                    }
                  }
                }}
                onKeyUp={(event) => handleEditorSelectionEvent('preview', event)}
                onPaste={(event) => void handlePaste(event)}
                onDrop={(event) => void handleDrop(event)}
                onDragOver={(event) => event.preventDefault()}
                spellCheck={false}
              />
            </article>
          </section>

          {isAiDrawerOpen && (
            <aside className="ai-drawer ai-drawer--dock">
              <header className="ai-drawer-header">
                <div className="ai-drawer-title-group">
                  <h2>AI 会话</h2>
                  {aiContextScope === 'section' && (
                    <div className="ai-section-indicator">
                      <span className="ai-section-indicator__label">当前章节</span>
                      <strong>{getOutlineItemById(effectiveAiSectionId)?.text ?? '未定位章节'}</strong>
                    </div>
                  )}
                </div>
                <div className="ai-drawer-header-actions">
                </div>
              </header>

              <div className="ai-toolbar">
                <div className="ai-toolbar-row ai-toolbar-row--fields">
                  <label className="translation-field ai-toolbar-field">
                    <span>上下文</span>
                    <select value={aiContextScope} onChange={(event) => setAiContextScope(event.target.value as AIContextScope)}>
                      <option value="selection">当前选区</option>
                      <option value="section">当前章节</option>
                      <option value="document">整篇文档</option>
                    </select>
                  </label>
                  <label className="translation-field ai-toolbar-field">
                    <span>模型档案</span>
                    <select value={activeAiProfileId} onChange={(event) => setActiveAiProfileId(event.target.value)}>
                      {aiProfiles.map((profile) => (
                        <option key={profile.id} value={profile.id}>
                          {profile.name}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>

                <div className="ai-toolbar-row ai-toolbar-row--actions">
                  <div className="ai-quick-actions ai-quick-actions--compact">
                    <button type="button" onClick={() => void handleAiAction('rewrite')} disabled={isAiLoading}>
                      改写
                    </button>
                    <button type="button" onClick={() => void handleAiAction('expand')} disabled={isAiLoading}>
                      扩写
                    </button>
                    <button type="button" onClick={() => void handleAiAction('shorten')} disabled={isAiLoading}>
                      精简
                    </button>
                    <button type="button" onClick={() => void handleAiAction('continue')} disabled={isAiLoading}>
                      续写
                    </button>
                  </div>
                  <div className="ai-profile-strip-actions">
                    <button type="button" onClick={() => setIsAiConfigOpen(true)}>
                      配置
                    </button>
                    <button type="button" onClick={() => void handleTestAiConnection()} disabled={isAiTesting}>
                      {isAiTesting ? '测试中...' : '测试'}
                    </button>
                  </div>
                </div>
              </div>

              <div className="ai-chat-list">
                {aiMessages.map((message) => (
                  <article key={message.id} className={`ai-chat-message ai-chat-message--${message.role}`}>
                    <header>
                      <strong>{message.role === 'user' ? '你' : 'AI 助手'}</strong>
                      <span>
                        {message.contextScope === 'selection'
                          ? '当前选区'
                          : message.contextScope === 'section'
                            ? '当前章节'
                            : '整篇文档'}
                      </span>
                    </header>
                    <pre>{message.content}</pre>
                    {message.role === 'assistant' && (
                      <footer className="ai-chat-actions">
                        <button
                          type="button"
                          onClick={() =>
                            insertIntoCurrentPosition(message.content, {
                              preferProgrammatic: true,
                              sectionId: message.sectionId ?? effectiveAiSectionId
                            })
                          }
                        >
                          插入到当前位置
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const sanitized = sanitizeAiInsertedMarkdown(message.content);
                            const restoredSelectionRange = restoreSavedEditorRange('selection');
                            if (insertMarkdownIntoRichEditor(sanitized, restoredSelectionRange, { recordHistory: true })) {
                              setStatusMessage('已用 AI 结果替换选区');
                              return;
                            }

                            const selection = getPreferredSelection();
                            const fallbackSelection = message.selectionRange;
                            const targetSelection = selection ?? (fallbackSelection
                              ? {
                                  start: fallbackSelection.start,
                                  end: fallbackSelection.end,
                                  text: doc.content.slice(fallbackSelection.start, fallbackSelection.end),
                                  source: 'preview' as const
                                }
                              : null);
                            if (!targetSelection) {
                              setStatusMessage('当前没有可替换的选区，请先选中文本');
                              return;
                            }
                            replaceContentRange(targetSelection.start, targetSelection.end, sanitized, targetSelection.source, {
                              start: targetSelection.start,
                              end: targetSelection.start + sanitized.length
                            });
                            setStatusMessage('已用 AI 结果替换选区');
                          }}
                        >
                          替换当前选区
                        </button>
                        <button
                          type="button"
                          onClick={() => appendToCurrentSection(message.content, message.sectionRange, message.sectionId)}
                        >
                          追加到当前章节
                        </button>
                        <button type="button" onClick={() => void copyText(message.content, '已复制 AI 回答')}>
                          复制
                        </button>
                      </footer>
                    )}
                  </article>
                ))}
                {isAiLoading && <div className="ai-chat-loading">AI 正在生成结果...</div>}
              </div>

              <footer className="ai-composer">
                <textarea
                  value={aiInput}
                  onChange={(event) => setAiInput(event.target.value)}
                  placeholder="例如：帮我总结当前章节、改写这段说明、补全后续步骤..."
                  onKeyDown={(event) => {
                    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
                      event.preventDefault();
                      void handleSendAiMessage();
                    }
                  }}
                />
                <div className="ai-composer-actions">
                  <span>按 Ctrl/Cmd + Enter 发送</span>
                  <button type="button" className="button-primary" onClick={() => void handleSendAiMessage()} disabled={isAiLoading || !aiInput.trim()}>
                    发送
                  </button>
                </div>
              </footer>
            </aside>
          )}
        </div>

        <footer className="status-bar">
          <div className="status-bar-group">
            <span>{stats.words} 词</span>
            <span>{stats.characters} 字符</span>
            <span>{stats.lines} 行</span>
          </div>
          <span className="status-bar-message">{statusMessage}</span>
        </footer>

        {contextMenu &&
          createPortal(
            <div
              className="context-menu"
              style={{ left: contextMenu.x, top: contextMenu.y }}
              onClick={(event) => event.stopPropagation()}
              onContextMenu={(event) => event.preventDefault()}
            >
              {contextMenuGroups.map((group, index) => (
                <div key={group.key || `group-${index}`} className="context-menu-group">
                <div className="context-menu-group-title">{group.title}</div>
                <div className={`context-menu-group-items columns-${group.columns ?? 2}`}>
                  {group.items.map((item) => (
                    <button
                      key={item.key}
                      type="button"
                      className={item.active ? 'is-active' : ''}
                      onClick={() => runContextAction(item.onClick)}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>
                </div>
              ))}
            </div>,
            document.body
          )}
        </main>

      {isAiConfigOpen && (
        <div className="ai-config-backdrop" onClick={() => setIsAiConfigOpen(false)}>
          <section className="ai-config-modal" onClick={(event) => event.stopPropagation()} aria-label="模型配置中心">
            <header className="ai-config-header">
              <div>
                <span className="card-label">模型配置中心</span>
                <h2>后台管理 AI 配置</h2>
                <p>支持同时保存多个厂商模型，前台只负责切换当前使用的档案。</p>
                <p className="ai-config-file-path">
                  配置文件：{aiConfigFilePath ?? '首次保存后将自动生成到应用目录中的“ai-profiles.jsonc”'}
                </p>
              </div>
              <button type="button" className="panel-toggle" onClick={() => setIsAiConfigOpen(false)}>
                关闭
              </button>
            </header>

            <div className="ai-config-layout">
              <aside className="ai-config-sidebar">
                <div className="ai-config-sidebar-actions">
                  <button type="button" onClick={() => addAiProfile('custom')}>
                    新增配置
                  </button>
                  <button type="button" onClick={duplicateActiveAiProfile}>
                    复制当前
                  </button>
                </div>

                <div className="ai-config-profile-list">
                  {aiProfiles.map((profile) => (
                    <button
                      key={profile.id}
                      type="button"
                      className={`ai-config-profile-item ${profile.id === activeAiProfileId ? 'is-active' : ''}`}
                      onClick={() => setActiveAiProfileId(profile.id)}
                    >
                      <strong>{profile.name}</strong>
                      <span>{profile.provider === 'builtin' ? '内置助手' : profile.model || '未配置模型'}</span>
                    </button>
                  ))}
                </div>
              </aside>

              <div className="ai-config-form">
                <div className="translation-settings ai-settings-grid">
                  <label className="translation-field">
                    <span>配置名称</span>
                    <input
                      type="text"
                      value={activeAiProfile.name}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          name: event.target.value || '未命名模型'
                        }))
                      }
                      placeholder="例如 DeepSeek 正式环境"
                    />
                  </label>
                  <label className="translation-field">
                    <span>模型预设</span>
                    <select value={currentAiPreset} onChange={(event) => applyAiPreset(event.target.value as AIPresetKey)}>
                      <option value="builtin">内置文档助手</option>
                      <option value="siliconflow">SiliconFlow</option>
                      <option value="dashscope">阿里云百炼</option>
                      <option value="deepseek">DeepSeek</option>
                      <option value="hunyuan">腾讯混元</option>
                      <option value="custom">自定义兼容接口</option>
                    </select>
                  </label>
                  <label className="translation-field">
                    <span>AI 提供方</span>
                    <select
                      value={aiSettings.provider}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          provider: event.target.value === 'openai-compatible' ? 'openai-compatible' : 'builtin'
                        }))
                      }
                    >
                      <option value="builtin">内置文档助手</option>
                      <option value="openai-compatible">OpenAI-compatible</option>
                    </select>
                  </label>
                  <label className="translation-field">
                    <span>API Base URL</span>
                    <input
                      type="text"
                      value={aiSettings.apiBase}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          apiBase: event.target.value
                        }))
                      }
                      placeholder="例如 https://api.deepseek.com/v1"
                    />
                  </label>
                  <label className="translation-field">
                    <span>API Key</span>
                    <input
                      type="password"
                      value={aiSettings.apiKey}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          apiKey: event.target.value
                        }))
                      }
                      placeholder="在这里保存当前配置的 API Key"
                    />
                  </label>
                  <label className="translation-field">
                    <span>模型名称</span>
                    <input
                      type="text"
                      value={aiSettings.model}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          model: event.target.value
                        }))
                      }
                      placeholder="例如 deepseek-chat"
                    />
                  </label>
                  <label className="translation-field">
                    <span>温度</span>
                    <input
                      type="number"
                      min="0"
                      max="1.2"
                      step="0.1"
                      value={aiSettings.temperature}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          temperature: Number.parseFloat(event.target.value) || 0
                        }))
                      }
                    />
                  </label>
                  <label className="translation-field">
                    <span>最大输出长度</span>
                    <input
                      type="number"
                      min="200"
                      max="4000"
                      step="100"
                      value={aiSettings.maxTokens}
                      onChange={(event) =>
                        updateActiveAiProfile((current) => ({
                          ...current,
                          maxTokens: Number.parseInt(event.target.value || '1200', 10)
                        }))
                      }
                    />
                  </label>
                </div>

                <p className="ai-preset-hint">{currentAiPresetConfig.description}</p>
                <div className="ai-preset-actions">
                  <button type="button" onClick={() => void openPresetConsole()} disabled={!currentAiPresetConfig.consoleUrl}>
                    获取 API Key
                  </button>
                  <button
                    type="button"
                    onClick={restoreAiPresetDefaults}
                    disabled={currentAiPreset === 'builtin' || currentAiPreset === 'custom'}
                  >
                    恢复推荐模型
                  </button>
                  <button type="button" onClick={() => void handleTestAiConnection()} disabled={isAiTesting}>
                    {isAiTesting ? '测试中...' : '测试当前配置'}
                  </button>
                  <button type="button" onClick={deleteActiveAiProfile} disabled={aiProfiles.length <= 1}>
                    删除当前配置
                  </button>
                </div>

                <div className="ai-log-panel">
                  <div className="ai-log-header">
                    <strong>AI 请求日志</strong>
                    <span>最近 {aiRequestLogs.length} 条</span>
                  </div>
                  <div className="ai-log-list">
                    {aiRequestLogs.length === 0 && <div className="ai-log-empty">尚无日志，先点一次“测试当前配置”或发起 AI 请求。</div>}
                    {aiRequestLogs.map((log) => (
                      <article key={log.id} className={`ai-log-entry is-${log.status}`}>
                        <header>
                          <strong>{log.time}</strong>
                          <span>
                            {log.kind} / {log.status}
                          </span>
                        </header>
                        <p>{log.detail}</p>
                        <code>{log.endpoint}</code>
                      </article>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </section>
        </div>
      )}

      {docxImportPreview && (
        <div className="beautify-modal-backdrop" onClick={() => setDocxImportPreview(null)}>
          <section className="beautify-modal docx-import-modal" onClick={(event) => event.stopPropagation()} aria-label="Word 导入预览">
            <header className="beautify-modal-header">
              <div>
                <span className="card-label">Word 导入预览</span>
                <h2>{docxImportPreview.sourceFileName}</h2>
                <p>先确认基础转换结果，再决定导入为新文档或插入当前文档。</p>
              </div>
              <button type="button" className="panel-toggle" onClick={() => setDocxImportPreview(null)}>
                关闭
              </button>
            </header>

            <div className="beautify-summary">
              <span className="beautify-badge">{docxImportPreview.summary.headingCount} 个标题</span>
              <span className="beautify-badge">{docxImportPreview.summary.listCount} 个列表项</span>
              <span className="beautify-badge">{docxImportPreview.summary.tableCount} 个表格</span>
              <span className="beautify-badge">{docxImportPreview.summary.imageCount} 张图片</span>
              <span className="beautify-badge">{docxImportPreview.summary.codeBlockCount} 个代码块</span>
            </div>

            {docxImportPreview.messages.length > 0 && (
              <div className="docx-import-messages">
                {docxImportPreview.messages.map((message) => (
                  <p key={message}>{message}</p>
                ))}
              </div>
            )}

            <div className="beautify-compare">
              <section className="beautify-pane">
                <div className="beautify-pane-title">转换后的 Markdown</div>
                <pre className="docx-import-markdown-preview">
                  <code>{docxImportPreview.markdown}</code>
                </pre>
              </section>

              <section className="beautify-pane">
                <div className="beautify-pane-title">渲染预览</div>
                <div
                  className="markdown-body docx-import-render-preview"
                  dangerouslySetInnerHTML={{ __html: docxImportPreviewHtml }}
                />
              </section>
            </div>

            <footer className="beautify-actions">
              <span>
                {docxImportPreview.imageAssets.length > 0
                  ? canInsertDocxImportIntoCurrentDocument
                    ? '图片会在确认导入时写入目标 Markdown 同级的 assets 目录。'
                    : '当前文档尚未保存。若要插入含图片的 Word 内容，请先保存当前 Markdown 文档。'
                  : '当前文档不包含图片资源，可直接导入。'}
              </span>
              <div className="beautify-actions-group">
                <button type="button" onClick={() => setDocxImportPreview(null)}>
                  取消
                </button>
                <button
                  type="button"
                  onClick={() => void insertDocxIntoCurrentDocument()}
                  disabled={!canInsertDocxImportIntoCurrentDocument}
                >
                  插入当前文档
                </button>
                <button type="button" className="button-primary" onClick={() => void importDocxAsNewDocument()}>
                  导入为新文档
                </button>
              </div>
            </footer>
          </section>
        </div>
      )}

      {aiActionPreview && (
        <div className="beautify-modal-backdrop" onClick={() => setAiActionPreview(null)}>
          <section
            className="beautify-modal"
            onClick={(event) => event.stopPropagation()}
            aria-label="AI 处理预览"
          >
            <header className="beautify-modal-header">
              <div>
                <span className="card-label">AI 处理预览</span>
                <h2>
                  {aiActionPreview.action === 'rewrite'
                    ? '改写预览'
                    : aiActionPreview.action === 'expand'
                      ? '扩写预览'
                      : aiActionPreview.action === 'shorten'
                        ? '精简预览'
                        : '续写预览'}
                </h2>
                <p>当前使用：{aiActionPreview.providerLabel}</p>
              </div>
              <button type="button" className="panel-toggle" onClick={() => setAiActionPreview(null)}>
                关闭
              </button>
            </header>

            <div className="beautify-compare">
              <section className="beautify-pane">
                <div className="beautify-pane-title">原文</div>
                <div className="beautify-diff">
                  {aiActionDiffRows.map((row, index) => (
                    <div key={`ai-left-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.left || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>

              <section className="beautify-pane">
                <div className="beautify-pane-title">AI 结果</div>
                <div className="beautify-diff">
                  {aiActionDiffRows.map((row, index) => (
                    <div key={`ai-right-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.right || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>
            </div>

            <footer className="beautify-actions">
              <span>{aiActionPreview.hasChanges ? '确认后再决定插入、替换或追加。' : '本次 AI 处理没有产生明显变化。'}</span>
              <div className="beautify-actions-group">
                <button type="button" onClick={() => applyAiActionPreview('copy')}>
                  复制
                </button>
                <button type="button" onClick={() => applyAiActionPreview('insert')}>
                  插入到当前位置
                </button>
                <button type="button" onClick={() => applyAiActionPreview('append')}>
                  追加到当前章节
                </button>
                <button
                  type="button"
                  className="button-primary"
                  disabled={!aiActionPreview.selectionRange}
                  onClick={() => applyAiActionPreview('replace')}
                >
                  替换当前选区
                </button>
              </div>
            </footer>
          </section>
        </div>
      )}

      {translationPreview && (
        <div className="beautify-modal-backdrop" onClick={() => setTranslationPreview(null)}>
          <section
            className="beautify-modal"
            onClick={(event) => event.stopPropagation()}
            aria-label="翻译预览"
          >
            <header className="beautify-modal-header">
              <div>
                <span className="card-label">翻译预览</span>
                <h2>{translationPreview.scope === 'selection' ? '翻译当前选中内容' : '翻译整篇文档'}</h2>
                <p>当前使用：{translationPreview.providerLabel}</p>
              </div>
              <button type="button" className="panel-toggle" onClick={() => setTranslationPreview(null)}>
                关闭
              </button>
            </header>

            <div className="beautify-strengths">
              <span className="beautify-strengths-label">翻译方向</span>
              <div className="beautify-strengths-group">
                <button
                  type="button"
                  className={translationDirection === 'zh-to-en' ? 'is-active' : ''}
                  onClick={() => void openTranslationPreview('zh-to-en')}
                  disabled={isTranslating}
                >
                  中译英
                </button>
                <button
                  type="button"
                  className={translationDirection === 'en-to-zh' ? 'is-active' : ''}
                  onClick={() => void openTranslationPreview('en-to-zh')}
                  disabled={isTranslating}
                >
                  英译中
                </button>
              </div>
            </div>

            <div className="ai-profile-strip">
              <label className="translation-field">
                <span>当前模型档案</span>
                <select value={activeAiProfileId} onChange={(event) => setActiveAiProfileId(event.target.value)}>
                  {aiProfiles.map((profile) => (
                    <option key={profile.id} value={profile.id}>
                      {profile.name}
                    </option>
                  ))}
                </select>
              </label>
              <div className="ai-profile-strip-actions">
                <button type="button" onClick={() => setIsAiConfigOpen(true)}>
                  配置模型
                </button>
                <button type="button" onClick={() => void handleTestAiConnection()} disabled={isAiTesting}>
                  {isAiTesting ? '测试中...' : '测试连接'}
                </button>
              </div>
            </div>
            <p className="ai-preset-hint">当前档案：{activeAiProfile.name}。翻译会复用这套模型配置。</p>

            <div className="beautify-compare">
              <section className="beautify-pane">
                <div className="beautify-pane-title">原文</div>
                <div className="beautify-diff">
                  {translationDiffRows.map((row, index) => (
                    <div key={`translation-left-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.left || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>

              <section className="beautify-pane">
                <div className="beautify-pane-title">译文</div>
                <div className="beautify-diff">
                  {translationDiffRows.map((row, index) => (
                    <div key={`translation-right-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.right || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>
            </div>

            <footer className="beautify-actions">
              <span>{isTranslating ? '正在生成翻译预览...' : '确认后才会覆盖当前文档或选中内容。'}</span>
              <div className="beautify-actions-group">
                <button type="button" onClick={() => void openTranslationPreview(translationDirection)} disabled={isTranslating}>
                  重新翻译
                </button>
                <button type="button" onClick={() => setTranslationPreview(null)}>
                  取消
                </button>
                <button
                  type="button"
                  className="button-primary"
                  onClick={applyTranslationPreview}
                  disabled={!translationPreview.hasChanges || isTranslating}
                >
                  {translationPreview.scope === 'selection' ? '覆盖选中内容' : '覆盖整篇文档'}
                </button>
              </div>
            </footer>
          </section>
        </div>
      )}

      {beautifyPreview && (
        <div className="beautify-modal-backdrop" onClick={() => setBeautifyPreview(null)}>
          <section
            className="beautify-modal"
            onClick={(event) => event.stopPropagation()}
            aria-label="AI 美化预览"
          >
            <header className="beautify-modal-header">
              <div>
                <span className="card-label">AI 美化预览</span>
                <h2>{beautifyPreview.scope === 'selection' ? '优化当前选中内容' : '优化整篇文档'}</h2>
                <p>{beautifyPreview.summary}</p>
              </div>
              <button type="button" className="panel-toggle" onClick={() => setBeautifyPreview(null)}>
                关闭
              </button>
            </header>

            <div className="beautify-strengths">
              <span className="beautify-strengths-label">美化强度</span>
              <div className="beautify-strengths-group">
                <button
                  type="button"
                  className={beautifyPreview.strength === 'light' ? 'is-active' : ''}
                  onClick={() => handleBeautifyStrengthChange('light')}
                >
                  轻度
                </button>
                <button
                  type="button"
                  className={beautifyPreview.strength === 'standard' ? 'is-active' : ''}
                  onClick={() => handleBeautifyStrengthChange('standard')}
                >
                  标准
                </button>
                <button
                  type="button"
                  className={beautifyPreview.strength === 'deep' ? 'is-active' : ''}
                  onClick={() => handleBeautifyStrengthChange('deep')}
                >
                  深度
                </button>
              </div>
            </div>

            <div className="beautify-summary">
              {beautifyPreview.changes.map((item) => (
                <span key={item} className="beautify-badge">
                  {item}
                </span>
              ))}
            </div>

            <div className="beautify-compare">
              <section className="beautify-pane">
                <div className="beautify-pane-title">原文</div>
                <div className="beautify-diff">
                  {beautifyDiffRows.map((row, index) => (
                    <div key={`left-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.left || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>

              <section className="beautify-pane">
                <div className="beautify-pane-title">优化后</div>
                <div className="beautify-diff">
                  {beautifyDiffRows.map((row, index) => (
                    <div key={`right-${index}`} className={`beautify-diff-row is-${row.type}`}>
                      <span className="beautify-line-number">{index + 1}</span>
                      <code>{row.right || ' '}</code>
                    </div>
                  ))}
                </div>
              </section>
            </div>

            <footer className="beautify-actions">
              <span>{beautifyPreview.hasChanges ? '确认后才会写回文档。' : '本次分析没有发现需要调整的地方。'}</span>
              <div className="beautify-actions-group">
                <button type="button" onClick={() => setBeautifyPreview(null)}>
                  取消
                </button>
                <button
                  type="button"
                  className="button-primary"
                  onClick={applyBeautifyPreview}
                  disabled={!beautifyPreview.hasChanges}
                >
                  {beautifyPreview.scope === 'selection' ? '应用到选中内容' : '应用到整篇文档'}
                </button>
              </div>
            </footer>
          </section>
        </div>
      )}
    </div>
  );
}

export default App;
