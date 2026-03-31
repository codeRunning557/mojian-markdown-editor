
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
import DOMPurify from 'dompurify';
import { marked } from 'marked';

type ViewMode = 'write' | 'split' | 'preview';
type SelectionSource = 'editor' | 'preview';
type BeautifyStrength = 'light' | 'standard' | 'deep';

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

type TranslationDirection = 'zh-to-en' | 'en-to-zh';

type TranslationSettings = {
  provider: 'builtin' | 'openai-compatible';
  apiBase: string;
  apiKey: string;
  model: string;
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

type ThemeName = 'paper' | 'mist' | 'slate' | 'graphite';

const recentFilesStorageKey = 'paperflow.recent-files';
const translationSettingsStorageKey = 'paperflow.translation-settings';
const themeStorageKey = 'paperflow.theme';
const allowedUriPattern = /^(?:(?:https?|file|mailto|tel|data):|[^a-z]|[a-z+.\-]+(?:[^a-z+.\-:]|$))/i;

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
  let headingIndex = 0;

  return rawHtml.replace(/<(h[1-6])>/g, (_match, tagName) => {
    const id = outline[headingIndex]?.id ?? `heading-${headingIndex}`;
    headingIndex += 1;
    return `<${tagName} data-outline-id="${id}">`;
  });
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

function loadTranslationSettings(): TranslationSettings {
  try {
    const raw = window.localStorage.getItem(translationSettingsStorageKey);
    if (!raw) {
      return {
        provider: 'builtin',
        apiBase: '',
        apiKey: '',
        model: ''
      };
    }

    const parsed = JSON.parse(raw) as Partial<TranslationSettings>;
    return {
      provider: parsed.provider === 'openai-compatible' ? 'openai-compatible' : 'builtin',
      apiBase: parsed.apiBase ?? '',
      apiKey: parsed.apiKey ?? '',
      model: parsed.model ?? ''
    };
  } catch {
    return {
      provider: 'builtin',
      apiBase: '',
      apiKey: '',
      model: ''
    };
  }
}

function loadThemePreference(): ThemeName {
  try {
    const raw = window.localStorage.getItem(themeStorageKey);
    if (raw === 'mist' || raw === 'slate' || raw === 'graphite') {
      return raw;
    }
  } catch {
    return 'paper';
  }

  return 'paper';
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
  settings: TranslationSettings,
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
      temperature: 0.1,
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
  settings: TranslationSettings
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
  const browserFileInputRef = useRef<HTMLInputElement | null>(null);
  const browserImageInputRef = useRef<HTMLInputElement | null>(null);
  const previewEditorRef = useRef<HTMLDivElement | null>(null);
  const lastRenderedMarkdownRef = useRef(starterDocument);
  const lastSelectionRef = useRef<{ start: number; end: number; source: SelectionSource } | null>(null);
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
  const [translationSettings, setTranslationSettings] = useState<TranslationSettings>(() => loadTranslationSettings());
  const [theme, setTheme] = useState<ThemeName>(() => loadThemePreference());
  const [translationPreview, setTranslationPreview] = useState<TranslationPreviewState | null>(null);
  const [isTranslating, setIsTranslating] = useState(false);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);

  const isDirty = doc.content !== lastSavedContent;
  const fileName = formatFileName(doc.filePath, doc.displayName);

  const stats = useMemo(() => {
    const trimmed = doc.content.trim();
    const words = trimmed ? trimmed.split(/\s+/).length : 0;
    const characters = doc.content.length;
    const lines = normalizeLineEndings(doc.content).split('\n').length;

    return { words, characters, lines };
  }, [doc.content]);

  const outline = useMemo<OutlineItem[]>(() => extractOutlineFromMarkdown(doc.content), [doc.content]);
  const editableHtml = useMemo(() => buildEditableHtml(doc.content, outline), [doc.content, outline]);

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
    const root = previewEditorRef.current;
    if (!root) {
      return;
    }

    if (document.activeElement === root && lastRenderedMarkdownRef.current === doc.content) {
      return;
    }

    root.innerHTML = editableHtml;
    lastRenderedMarkdownRef.current = doc.content;
  }, [doc.content, editableHtml]);

  useEffect(() => {
    window.localStorage.setItem(recentFilesStorageKey, JSON.stringify(recentFiles));
  }, [recentFiles]);

  useEffect(() => {
    window.localStorage.setItem(translationSettingsStorageKey, JSON.stringify(translationSettings));
  }, [translationSettings]);

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
    if (!desktopApi || !doc.filePath || !isDirty) {
      return;
    }

    const timeoutId = window.setTimeout(async () => {
      try {
        await desktopApi.saveMarkdownFile({
          filePath: doc.filePath,
          content: doc.content,
          forceDialog: false
        });
        setLastSavedContent(doc.content);
        setStatusMessage(`已自动保存 ${fileName}`);
      } catch (error) {
        const message = error instanceof Error ? error.message : '未知自动保存错误';
        setStatusMessage(`自动保存失败：${message}`);
      }
    }, 900);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [desktopApi, doc.content, doc.filePath, fileName, isDirty]);

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

      if (action === 'save-file') {
        await handleSaveFile(false);
      }

      if (action === 'save-file-as') {
        await handleSaveFile(true);
      }

      if (action === 'reference-image') {
        await handleReferenceImage();
      }

      if (action === 'ai-beautify') {
        openBeautifyPreview();
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
    });

    return () => {
      dispose();
    };
  }, [desktopApi, doc.content, doc.filePath, isDirty]);

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

  function getPreferredSelection() {
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

  function syncRichEditorToDoc() {
    const root = previewEditorRef.current;
    if (!root) {
      return;
    }

    const markdown = serializeRichEditorContent(root);
    lastRenderedMarkdownRef.current = markdown;
    setDoc((current) => (current.content === markdown ? current : { ...current, content: markdown }));
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

  function insertMarkdownIntoRichEditor(markdown: string) {
    const root = previewEditorRef.current;
    const selection = window.getSelection();

    if (!root || !selection || selection.rangeCount === 0) {
      return false;
    }

    const range = selection.getRangeAt(0);
    if (!root.contains(range.commonAncestorContainer)) {
      return false;
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

  function insertIntoCurrentPosition(markdown: string) {
    if (insertMarkdownIntoRichEditor(markdown)) {
      return;
    }

    const selection = getPreferredSelection();

    if (!selection) {
      const appendPrefix = doc.content.endsWith('\n') ? '' : '\n';
      const nextContent = `${doc.content}${appendPrefix}${markdown}\n`;
      const cursor = nextContent.length;
      setDoc((current) => ({
        ...current,
        content: nextContent
      }));
      lastSelectionRef.current = {
        start: cursor,
        end: cursor,
        source: 'preview'
      };
      focusEditorSoon('preview', cursor);
      return;
    }

    replaceContentRange(selection.start, selection.end, markdown, selection.source);
  }

  function insertHeadingAtCurrentPosition(level: 1 | 2 | 3) {
    const prefix = `${'#'.repeat(level)} `;
    const fallbackText = level === 1 ? '一级标题' : level === 2 ? '二级标题' : '三级标题';
    const richSelectionText = getRichEditorTextSelection();

    if (insertMarkdownIntoRichEditor(`${prefix}${richSelectionText || fallbackText}`)) {
      setStatusMessage(`已插入 ${fallbackText}`);
      return;
    }

    const selection = getPreferredSelection();

    if (!selection) {
      const appendPrefix = doc.content.endsWith('\n') ? '' : '\n\n';
      const snippet = `${appendPrefix}${prefix}${fallbackText}\n\n`;
      const start = doc.content.length + appendPrefix.length + prefix.length;
      const end = start + fallbackText.length;
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

    if (insertMarkdownIntoRichEditor(`\`\`\`\n${richSelectionText || '在此输入代码'}\n\`\`\``)) {
      setStatusMessage('已插入代码块');
      return;
    }

    const selection = getPreferredSelection();

    if (!selection) {
      const appendPrefix = doc.content.endsWith('\n') ? '' : '\n\n';
      const snippet = `${appendPrefix}\`\`\`\n在此输入代码\n\`\`\`\n`;
      const start = doc.content.length + appendPrefix.length + 4;
      const end = start + '在此输入代码'.length;
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
    setDoc({
      filePath: null,
      displayName: '未命名.md',
      content: '# 未命名文档\n\n'
    });
    setLastSavedContent('');
    setViewMode('preview');
    setBeautifyPreview(null);
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

      setDoc({
        filePath: nextPath,
        displayName: nextName,
        content: nextContent
      });
      setLastSavedContent(nextContent);
      setViewMode('preview');
      setBeautifyPreview(null);
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

      setDoc({
        filePath,
        displayName: file.name,
        content
      });
      setLastSavedContent(content);
      setViewMode('preview');
      setBeautifyPreview(null);
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
    setContextMenu({
      x: event.clientX,
      y: event.clientY
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
      graphite: '石墨主题'
    } satisfies Record<ThemeName, string>;
    setStatusMessage(`已切换到${labels[nextTheme]}`);
  }

  async function openTranslationPreview(direction = translationDirection) {
    const selection = getPreferredSelection();
    const useSelection = Boolean(selection && selection.text.trim());
    const sourceText = useSelection ? selection!.text : doc.content;

    setTranslationDirection(direction);
    setIsTranslating(true);
    try {
      const result = await translateDocumentContent(sourceText, direction, translationSettings);
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
      setStatusMessage(useSelection ? '已生成选中内容的翻译预览' : '已生成整篇文档的翻译预览');
    } catch (error) {
      const message = error instanceof Error ? error.message : '未知翻译错误';
      setStatusMessage(`翻译失败：${message}`);
    } finally {
      setIsTranslating(false);
    }
  }

  function applyTranslationPreview() {
    if (!translationPreview) {
      return;
    }

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
    lastSelectionRef.current = null;
    syncRichEditorToDoc();
  }

  function handleEditorSelectionEvent(
    source: SelectionSource,
    _event:
      | ReactMouseEvent<HTMLDivElement>
      | KeyboardEvent<HTMLDivElement>
      | FormEvent<HTMLDivElement>
  ) {
    lastSelectionRef.current = null;
  }

  const contextMenuGroups = [
    [
      { key: 'ctx-h1', label: '插入一级标题', onClick: () => insertHeadingAtCurrentPosition(1) },
      { key: 'ctx-h2', label: '插入二级标题', onClick: () => insertHeadingAtCurrentPosition(2) },
      { key: 'ctx-h3', label: '插入三级标题', onClick: () => insertHeadingAtCurrentPosition(3) },
      { key: 'ctx-code', label: '插入代码块', onClick: () => insertCodeBlockAtCurrentPosition() }
    ],
    [
      { key: 'ctx-ref-image', label: '引用图片', onClick: () => handleReferenceImage() },
      { key: 'ctx-ai', label: 'AI 美化', onClick: () => openBeautifyPreview() },
      { key: 'ctx-translate', label: '翻译', onClick: () => openTranslationPreview() }
    ],
    [
      { key: 'ctx-theme-paper', label: '米白主题', onClick: () => applyTheme('paper'), active: theme === 'paper' },
      { key: 'ctx-theme-mist', label: '雾蓝主题', onClick: () => applyTheme('mist'), active: theme === 'mist' },
      { key: 'ctx-theme-slate', label: '灰砚主题', onClick: () => applyTheme('slate'), active: theme === 'slate' },
      { key: 'ctx-theme-graphite', label: '石墨主题', onClick: () => applyTheme('graphite'), active: theme === 'graphite' }
    ]
  ];

  return (
    <div className="shell" data-theme={theme}>
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
              className={`sidebar-link sidebar-link--outline level-${item.level}`}
              onClick={() => jumpToHeading(item.id)}
            >
              <strong>{item.text}</strong>
            </button>
          ))}
        </div>
      </aside>

      <main className="workspace">
        <section className="editor-grid editor-grid--preview" onContextMenu={handleContextMenu}>
          <article className="panel panel--workspace">
            <div className="panel-heading">
              <span className="panel-title">实时预览编辑</span>
              <span className="panel-meta">单栏编辑工作区</span>
            </div>

            <div
              ref={previewEditorRef}
              className="markdown-body rich-editor preview-editor--workspace preview-editor--single"
              contentEditable
              suppressContentEditableWarning
              data-placeholder="在这里直接编辑富文本 Markdown 内容..."
              onInput={handleRichEditorInput}
              onClick={(event) => handleEditorSelectionEvent('preview', event)}
              onKeyUp={(event) => handleEditorSelectionEvent('preview', event)}
              onPaste={(event) => void handlePaste(event)}
              onDrop={(event) => void handleDrop(event)}
              onDragOver={(event) => event.preventDefault()}
              spellCheck={false}
            />
          </article>
        </section>

        <footer className="status-bar">
          <div className="status-bar-group">
            <span>{stats.words} 词</span>
            <span>{stats.characters} 字符</span>
            <span>{stats.lines} 行</span>
          </div>
          <span className="status-bar-message">{statusMessage}</span>
        </footer>

        {contextMenu && (
          <div
            className="context-menu"
            style={{ left: contextMenu.x, top: contextMenu.y }}
            onClick={(event) => event.stopPropagation()}
          >
            {contextMenuGroups.map((group, index) => (
              <div key={`group-${index}`} className="context-menu-group">
                {group.map((item) => (
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
            ))}
          </div>
        )}
      </main>

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

            <div className="translation-settings">
              <label className="translation-field">
                <span>翻译提供方</span>
                <select
                  value={translationSettings.provider}
                  onChange={(event) =>
                    setTranslationSettings((current) => ({
                      ...current,
                      provider: event.target.value === 'openai-compatible' ? 'openai-compatible' : 'builtin'
                    }))
                  }
                >
                  <option value="builtin">内置轻量翻译器</option>
                  <option value="openai-compatible">自定义兼容 API</option>
                </select>
              </label>
              <label className="translation-field">
                <span>API Base URL</span>
                <input
                  type="text"
                  value={translationSettings.apiBase}
                  onChange={(event) =>
                    setTranslationSettings((current) => ({
                      ...current,
                      apiBase: event.target.value
                    }))
                  }
                  placeholder="例如 https://api.openai.com/v1"
                />
              </label>
              <label className="translation-field">
                <span>API Key</span>
                <input
                  type="password"
                  value={translationSettings.apiKey}
                  onChange={(event) =>
                    setTranslationSettings((current) => ({
                      ...current,
                      apiKey: event.target.value
                    }))
                  }
                  placeholder="留空则使用内置翻译器"
                />
              </label>
              <label className="translation-field">
                <span>模型名称</span>
                <input
                  type="text"
                  value={translationSettings.model}
                  onChange={(event) =>
                    setTranslationSettings((current) => ({
                      ...current,
                      model: event.target.value
                    }))
                  }
                  placeholder="例如 gpt-4.1-mini"
                />
              </label>
            </div>

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
