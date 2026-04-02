const path = require('node:path');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const { app, BrowserWindow, dialog, ipcMain, Menu, shell } = require('electron');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const { gfm } = require('turndown-plugin-gfm');

let mainWindow = null;
let allowWindowClose = false;
let isHandlingClose = false;
let latestDocumentState = {
  filePath: null,
  displayName: '未命名.md',
  content: '',
  isDirty: false
};

process.on('uncaughtException', (error) => {
  console.error('[main] uncaught exception');
  console.error(error);
  void emergencySaveDocument('uncaught-exception');
});

function emitMenuAction(action) {
  if (!mainWindow) {
    return;
  }

  mainWindow.webContents.send('menu-action', action);
}

function getRecoveryDirectory() {
  return path.join(app.getPath('documents'), '墨笺自动恢复');
}

function getAiConfigDirectory() {
  return app.isPackaged ? path.dirname(process.execPath) : app.getAppPath();
}

function getAiConfigLocalFilePath() {
  return path.join(app.getPath('userData'), 'ai-profiles.local.json');
}

let resolvedAiConfigTemplateFilePath = null;
const aiConfigTemplateFileName = 'ai-profiles.jsonc';
const legacyAiConfigTemplateFileName = 'ai-profiles.json';

async function canWriteToDirectory(directoryPath) {
  try {
    await fs.mkdir(directoryPath, { recursive: true });
    const probePath = path.join(directoryPath, `.mojian-write-test-${process.pid}-${Date.now()}.tmp`);
    await fs.writeFile(probePath, '', 'utf8');
    await fs.unlink(probePath);
    return true;
  } catch {
    return false;
  }
}

async function resolveAiConfigFilePath() {
  if (resolvedAiConfigTemplateFilePath) {
    return resolvedAiConfigTemplateFilePath;
  }

  const preferredDirectory = getAiConfigDirectory();
  if (await canWriteToDirectory(preferredDirectory)) {
    resolvedAiConfigTemplateFilePath = path.join(preferredDirectory, aiConfigTemplateFileName);
    return resolvedAiConfigTemplateFilePath;
  }

  const fallbackDirectory = path.join(app.getPath('userData'), 'config');
  await fs.mkdir(fallbackDirectory, { recursive: true });
  resolvedAiConfigTemplateFilePath = path.join(fallbackDirectory, aiConfigTemplateFileName);
  return resolvedAiConfigTemplateFilePath;
}

function getLegacyAiConfigFilePath(templateFilePath) {
  return path.join(path.dirname(templateFilePath), legacyAiConfigTemplateFileName);
}

function normalizePersistedAiProfile(profile) {
  const name = String(profile?.name || '').trim() || '未命名模型';
  return {
    name,
    provider: profile?.provider === 'openai-compatible' ? 'openai-compatible' : 'builtin',
    apiBase: String(profile?.apiBase || ''),
    apiKey: String(profile?.apiKey || ''),
    model: String(profile?.model || ''),
    temperature:
      typeof profile?.temperature === 'number' && Number.isFinite(profile.temperature)
        ? Math.min(1.2, Math.max(0, profile.temperature))
        : 0.2,
    maxTokens:
      typeof profile?.maxTokens === 'number' && Number.isFinite(profile.maxTokens)
        ? Math.min(4000, Math.max(200, Math.round(profile.maxTokens)))
        : 1200
  };
}

function mergePersistedAiProfilesByName(profiles) {
  const profileMap = new Map();

  for (const profile of Array.isArray(profiles) ? profiles : []) {
    const normalized = normalizePersistedAiProfile(profile);
    if (profileMap.has(normalized.name)) {
      profileMap.delete(normalized.name);
    }
    profileMap.set(normalized.name, normalized);
  }

  return Array.from(profileMap.values());
}

async function readJsonIfExists(filePath) {
  try {
    const raw = await fs.readFile(filePath, 'utf8');
    try {
      return JSON.parse(raw);
    } catch {
      return JSON.parse(stripJsonComments(raw));
    }
  } catch (error) {
    if (error && typeof error === 'object' && error.code === 'ENOENT') {
      return null;
    }

    throw error;
  }
}

function stripJsonComments(raw) {
  let result = '';
  let isInString = false;
  let isEscaped = false;
  let isInLineComment = false;
  let isInBlockComment = false;

  for (let index = 0; index < raw.length; index += 1) {
    const currentChar = raw[index];
    const nextChar = raw[index + 1];

    if (isInLineComment) {
      if (currentChar === '\n') {
        isInLineComment = false;
        result += currentChar;
      }
      continue;
    }

    if (isInBlockComment) {
      if (currentChar === '*' && nextChar === '/') {
        isInBlockComment = false;
        index += 1;
      }
      continue;
    }

    if (isInString) {
      result += currentChar;
      if (isEscaped) {
        isEscaped = false;
      } else if (currentChar === '\\') {
        isEscaped = true;
      } else if (currentChar === '"') {
        isInString = false;
      }
      continue;
    }

    if (currentChar === '/' && nextChar === '/') {
      isInLineComment = true;
      index += 1;
      continue;
    }

    if (currentChar === '/' && nextChar === '*') {
      isInBlockComment = true;
      index += 1;
      continue;
    }

    result += currentChar;
    if (currentChar === '"') {
      isInString = true;
      isEscaped = false;
    }
  }

  return result;
}

function buildAiConfigTemplateDocumentText() {
  const updatedAt = new Date().toISOString();
  return `{
  // 墨笺 AI 配置模板（JSONC）
  // 此文件仅用于展示可配置项及字段说明，不保存任何真实配置值。
  // 真实的模型地址、API Key、模型名等敏感信息会保存在当前系统用户的私有配置中，不会随应用包传播。
  "version": 1,

  // 模板更新时间，便于确认当前模板是否为最新结构。
  "updatedAt": "${updatedAt}",

  // 当前模板用途说明。
  "note": "此文件仅保留配置项字段与注释说明，不保存任何配置值。",

  "profileTemplate": {
    // 配置名称：用于区分不同厂商或不同用途的模型档案。
    "name": "",

    // AI 提供方：builtin 表示内置助手，openai-compatible 表示兼容 OpenAI 接口的模型服务。
    "provider": "",

    // 接口地址：通常填写厂商提供的 Base URL，例如 https://api.deepseek.com/v1。
    "apiBase": "",

    // API 密钥：真实使用时由用户在前台填写并私有保存，不会写入本模板文件。
    "apiKey": "",

    // 模型名称：例如 deepseek-chat、qwen-plus、gpt-4.1-mini。
    "model": "",

    // 温度：控制生成结果的发散程度，推荐范围 0.0 ~ 1.2。
    "temperature": "",

    // 最大输出 Token：限制单次回复的最大生成长度。
    "maxTokens": ""
  }
}
`;
}

async function removeLegacyAiConfigFile(templateFilePath) {
  const legacyFilePath = getLegacyAiConfigFilePath(templateFilePath);
  if (legacyFilePath === templateFilePath) {
    return;
  }

  try {
    await fs.unlink(legacyFilePath);
  } catch (error) {
    if (!(error && typeof error === 'object' && error.code === 'ENOENT')) {
      throw error;
    }
  }
}

async function ensureAiConfigTemplateFile() {
  const filePath = await resolveAiConfigFilePath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, buildAiConfigTemplateDocumentText(), 'utf8');
  await removeLegacyAiConfigFile(filePath);
  return filePath;
}

async function writeAiLocalConfigDocument(payload) {
  const filePath = getAiConfigLocalFilePath();
  await fs.mkdir(path.dirname(filePath), { recursive: true });

  const profiles = mergePersistedAiProfilesByName(payload?.profiles);
  const requestedActiveName = typeof payload?.activeProfileName === 'string' ? payload.activeProfileName.trim() : '';
  const activeProfileName =
    requestedActiveName && profiles.some((profile) => profile.name === requestedActiveName)
      ? requestedActiveName
      : profiles[0]?.name ?? null;

  const document = {
    version: 1,
    updatedAt: new Date().toISOString(),
    activeProfileName,
    profiles
  };

  await fs.writeFile(filePath, JSON.stringify(document, null, 2), 'utf8');
  return { filePath, activeProfileName, profiles };
}

async function readAiConfigDocument() {
  const templateFilePath = await resolveAiConfigFilePath();
  const localFilePath = getAiConfigLocalFilePath();
  const templateDocument = await readJsonIfExists(templateFilePath);
  const legacyTemplateDocument = templateDocument ?? (await readJsonIfExists(getLegacyAiConfigFilePath(templateFilePath)));
  const localDocument = await readJsonIfExists(localFilePath);

  if (localDocument) {
    await ensureAiConfigTemplateFile();
    const profiles = mergePersistedAiProfilesByName(localDocument?.profiles);
    const activeProfileName =
      typeof localDocument?.activeProfileName === 'string' &&
      profiles.some((profile) => profile.name === localDocument.activeProfileName)
        ? localDocument.activeProfileName
        : profiles[0]?.name ?? null;

    return { filePath: templateFilePath, activeProfileName, profiles };
  }

  const migratedProfiles = mergePersistedAiProfilesByName(legacyTemplateDocument?.profiles);
  const hasLegacyValues = Array.isArray(legacyTemplateDocument?.profiles) && legacyTemplateDocument.profiles.length > 0;

  if (hasLegacyValues) {
    const migrated = await writeAiLocalConfigDocument({
      activeProfileName: legacyTemplateDocument?.activeProfileName ?? null,
      profiles: migratedProfiles
    });
    await ensureAiConfigTemplateFile();
    return {
      filePath: templateFilePath,
      activeProfileName: migrated.activeProfileName,
      profiles: migrated.profiles
    };
  }

  await ensureAiConfigTemplateFile();
  return { filePath: templateFilePath, activeProfileName: null, profiles: [] };
}

async function writeAiConfigDocument(payload) {
  const writtenLocalDocument = await writeAiLocalConfigDocument(payload);
  const templateFilePath = await ensureAiConfigTemplateFile();
  return {
    filePath: templateFilePath,
    activeProfileName: writtenLocalDocument.activeProfileName,
    profiles: writtenLocalDocument.profiles
  };
}

async function persistMarkdownContent(targetWindow, payload) {
  const shouldShowDialog = !payload?.filePath || payload?.forceDialog;
  let filePath = payload?.filePath ?? '';

  if (shouldShowDialog) {
    const result = await dialog.showSaveDialog(targetWindow, {
      title: '保存 Markdown 文件',
      defaultPath: filePath || payload?.defaultPath || 'untitled.md',
      filters: [
        { name: 'Markdown', extensions: ['md'] },
        { name: '文本', extensions: ['txt'] }
      ]
    });

    if (result.canceled || !result.filePath) {
      return { canceled: true };
    }

    filePath = result.filePath;
  }

  await fs.writeFile(filePath, payload?.content ?? '', 'utf8');
  return { canceled: false, filePath };
}

async function emergencySaveDocument(reason) {
  if (!latestDocumentState.isDirty || !latestDocumentState.content) {
    return null;
  }

  try {
    if (latestDocumentState.filePath) {
      await fs.writeFile(latestDocumentState.filePath, latestDocumentState.content, 'utf8');
      console.log('[main] emergency saved document', { reason, filePath: latestDocumentState.filePath });
      latestDocumentState.isDirty = false;
      return latestDocumentState.filePath;
    }

    const recoveryDir = getRecoveryDirectory();
    await fs.mkdir(recoveryDir, { recursive: true });
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const safeName = sanitizeFileComponent(path.parse(latestDocumentState.displayName || '未命名文档').name, '未命名文档');
    const recoveryPath = path.join(recoveryDir, `${safeName}-${timestamp}.md`);
    await fs.writeFile(recoveryPath, latestDocumentState.content, 'utf8');
    console.log('[main] emergency saved recovery document', { reason, filePath: recoveryPath });
    latestDocumentState.isDirty = false;
    return recoveryPath;
  } catch (error) {
    console.error('[main] emergency save failed', reason, error);
    return null;
  }
}

function sanitizeFileComponent(value, fallback = 'asset') {
  const nextValue = String(value || '')
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '-')
    .trim();

  return nextValue || fallback;
}

function getDocxImageExtension(contentType) {
  const extensionMap = {
    'image/png': '.png',
    'image/jpeg': '.jpg',
    'image/gif': '.gif',
    'image/webp': '.webp',
    'image/svg+xml': '.svg',
    'image/bmp': '.bmp',
    'image/tiff': '.tiff'
  };

  return extensionMap[contentType] ?? '.png';
}

function normalizeMarkdownForImport(markdown) {
  const normalized = String(markdown || '')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return normalized ? `${normalized}\n` : '';
}

function createDocxTurndownService() {
  const service = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '*'
  });

  service.use(gfm);
  service.remove(['style', 'script']);
  service.keep(['sub', 'sup']);

  service.addRule('docxTable', {
    filter: 'table',
    replacement(_content, node) {
      return `\n\n${convertTableNodeToMarkdown(node)}\n\n`;
    }
  });

  service.addRule('docxImageToken', {
    filter: 'img',
    replacement(_content, node) {
      const src = node.getAttribute('src') ?? '';
      const alt = node.getAttribute('alt') ?? 'image';
      return src ? `![${alt}](${src})` : '';
    }
  });

  return service;
}

function buildDocxImportSummary(markdown, html, imageAssets) {
  return {
    headingCount: (markdown.match(/^#{1,6}\s+/gm) || []).length,
    listCount: (markdown.match(/^\s*(?:[-*+]|\d+\.)\s+/gm) || []).length,
    tableCount: (html.match(/<table[\s>]/gi) || []).length,
    imageCount: imageAssets.length,
    codeBlockCount: (markdown.match(/^```[\w-]*\n[\s\S]*?^```/gm) || []).length
  };
}

function escapeMarkdownTableCell(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\n+/g, '<br>')
    .replace(/\|/g, '\\|')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractTableCellText(cellNode) {
  const pieces = [];

  for (const childNode of Array.from(cellNode.childNodes || [])) {
    const childName = childNode.nodeName?.toLowerCase?.() ?? '';
    if (childName === 'p' || childName === 'div') {
      const text = String(childNode.textContent || '').trim();
      if (text) {
        pieces.push(text);
      }
      continue;
    }

    const text = String(childNode.textContent || '').trim();
    if (text) {
      pieces.push(text);
    }
  }

  return escapeMarkdownTableCell(pieces.join('<br>'));
}

function convertTableNodeToMarkdown(tableNode) {
  const rowNodes = Array.from(tableNode.querySelectorAll('tr'));
  if (rowNodes.length === 0) {
    return '';
  }

  const rows = rowNodes.map((rowNode) =>
    Array.from(rowNode.querySelectorAll('th, td')).map((cellNode) => extractTableCellText(cellNode))
  );

  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const normalizedRows = rows.map((row) => {
    const padded = [...row];
    while (padded.length < columnCount) {
      padded.push('');
    }
    return padded;
  });

  const header = normalizedRows[0];
  const separator = header.map(() => '---');
  const bodyRows = normalizedRows.slice(1);
  const markdownRows = [header, separator, ...bodyRows];

  return markdownRows.map((row) => `| ${row.join(' | ')} |`).join('\n');
}

function isLikelyCodeLine(line) {
  const trimmed = line.trim();
  if (!trimmed) {
    return false;
  }

  if (/^#{1,6}\s/.test(trimmed) || /^[-*+]\s/.test(trimmed)) {
    return false;
  }

  if (/^(npm|pnpm|yarn|node|npx|git|docker|docker-compose|kubectl|helm|curl|wget|ssh|scp|cd|ls|pwd|mkdir|rm|cp|mv|cat|echo|pip|python|python3|java|javac|mvn|gradle|go|cargo|rustc|mysql|psql|redis-cli|systemctl|service)\b/i.test(trimmed)) {
    return true;
  }

  if (/^(Get-|Set-|New-|Remove-|Write-|Start-|Stop-|Invoke-|Test-|Copy-|Move-)[A-Za-z]/.test(trimmed)) {
    return true;
  }

  if (/^(\$|#|\/\/|--|PS\s|C:\\|\/|\.\.?\/)/.test(trimmed)) {
    return true;
  }

  if (/^(import |from |def |class |if __name__ == ['"]__main__['"]|public |private |protected |interface |package |const |let |var |function )/.test(trimmed)) {
    return true;
  }

  if ((/[{}();=<>\[\]]/.test(trimmed) || /\s--[\w-]+/.test(trimmed)) && !/[\u4e00-\u9fff]{3,}/.test(trimmed)) {
    return true;
  }

  return false;
}

function inferCodeFenceLanguage(lines) {
  const sample = lines.join('\n');

  if (/^(Get-|Set-|New-|Remove-|Write-|Start-|Stop-|Invoke-|Test-|Copy-|Move-)/m.test(sample) || /\$env:|Write-Host|powershell/i.test(sample)) {
    return 'powershell';
  }

  if (/^(docker|docker-compose|npm|pnpm|yarn|git|kubectl|helm|curl|wget|cd|ls|pwd|mkdir|rm|cp|mv|cat|echo)\b/im.test(sample)) {
    return 'bash';
  }

  if (/^\s*(import |from |def |class |print\(|if __name__ == ['"]__main__['"])/m.test(sample)) {
    return 'python';
  }

  if (/^\s*(public |private |protected |class |interface |package |import java\.)/m.test(sample)) {
    return 'java';
  }

  if (/^\s*(const |let |var |function |import .* from |export )/m.test(sample)) {
    return 'javascript';
  }

  if (/^\s*\{[\s\S]*\}\s*$/.test(sample)) {
    return 'json';
  }

  if (/^\s*[\w-]+\s*:\s+\S+/m.test(sample) && !/[{}();]/.test(sample)) {
    return 'yaml';
  }

  if (/^\s*(select|insert|update|delete|create table|alter table)\b/im.test(sample)) {
    return 'sql';
  }

  return 'text';
}

function convertCodeLikeParagraphGroups(markdown) {
  const blocks = String(markdown || '').split(/\n{2,}/);
  const nextBlocks = [];

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const trimmed = block.trim();

    if (!trimmed || /^```/.test(trimmed) || /^\|/.test(trimmed) || /^</.test(trimmed)) {
      nextBlocks.push(trimmed);
      continue;
    }

    const group = [];
    let cursor = index;

    while (cursor < blocks.length) {
      const currentBlock = blocks[cursor].trim();
      if (!currentBlock || /^```/.test(currentBlock) || /^\|/.test(currentBlock) || /^</.test(currentBlock)) {
        break;
      }

      const lines = currentBlock
        .split('\n')
        .map((line) => line.trim())
        .filter(Boolean);

      if (lines.length === 0 || !lines.every(isLikelyCodeLine)) {
        break;
      }

      group.push(...lines);
      cursor += 1;
    }

    if (group.length >= 2) {
      const language = inferCodeFenceLanguage(group);
      nextBlocks.push(`\`\`\`${language === 'text' ? '' : language}\n${group.join('\n')}\n\`\`\``.trim());
      index = cursor - 1;
      continue;
    }

    nextBlocks.push(trimmed);
  }

  return nextBlocks.filter(Boolean).join('\n\n');
}

function postProcessImportedMarkdown(markdown) {
  return normalizeMarkdownForImport(convertCodeLikeParagraphGroups(markdown));
}

async function createUniqueFilePath(targetDir, fileName) {
  const parsed = path.parse(fileName);
  const baseName = sanitizeFileComponent(parsed.name, 'asset');
  const extension = parsed.ext || '.png';

  let candidateName = `${baseName}${extension}`;
  let counter = 1;

  while (fsSync.existsSync(path.join(targetDir, candidateName))) {
    candidateName = `${baseName}-${counter}${extension}`;
    counter += 1;
  }

  return path.join(targetDir, candidateName);
}

async function materializeImportedMarkdown(markdown, imageAssets, documentPath) {
  let nextMarkdown = normalizeMarkdownForImport(markdown);

  if (!imageAssets?.length) {
    return nextMarkdown;
  }

  const targetDir = path.join(path.dirname(documentPath), 'assets');
  await fs.mkdir(targetDir, { recursive: true });

  for (const asset of imageAssets) {
    const suggestedName = sanitizeFileComponent(path.parse(asset.originalName || '').name, `docx-image-${Date.now()}`);
    const extension = path.extname(asset.originalName || '') || getDocxImageExtension(asset.contentType);
    const assetPath = await createUniqueFilePath(targetDir, `${suggestedName}${extension}`);
    const base64Data = String(asset.dataUrl || '').replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '');
    const buffer = Buffer.from(base64Data, 'base64');
    await fs.writeFile(assetPath, buffer);

    const relativePath = path.relative(path.dirname(documentPath), assetPath).replace(/\\/g, '/');
    const markdownPath = relativePath.startsWith('.') ? relativePath : `./${relativePath}`;
    nextMarkdown = nextMarkdown.split(asset.token).join(markdownPath);
  }

  return nextMarkdown;
}

async function convertDocxFileToPreview(filePath) {
  const imageAssets = [];
  const result = await mammoth.convertToHtml(
    { path: filePath },
    {
      ignoreEmptyParagraphs: false,
      convertImage: mammoth.images.imgElement(async (image) => {
        const base64 = await image.readAsBase64String();
        const extension = getDocxImageExtension(image.contentType);
        const token = `__DOCX_IMAGE_${imageAssets.length}__`;
        const originalName = `docx-image-${imageAssets.length + 1}${extension}`;
        const dataUrl = `data:${image.contentType};base64,${base64}`;

        imageAssets.push({
          token,
          originalName,
          contentType: image.contentType,
          dataUrl
        });

        return {
          src: token,
          alt: path.parse(originalName).name
        };
      })
    }
  );

  const html = String(result.value || '');
  const turndownService = createDocxTurndownService();
  const markdown = postProcessImportedMarkdown(turndownService.turndown(html));
  const previewHtml = imageAssets.reduce((currentHtml, asset) => currentHtml.split(asset.token).join(asset.dataUrl), html);
  const summary = buildDocxImportSummary(markdown, html, imageAssets);
  const messages = result.messages.map((message) => `${message.type === 'warning' ? '提示' : '消息'}：${message.message}`);

  return {
    sourceFilePath: filePath,
    sourceFileName: path.basename(filePath),
    markdown,
    previewHtml,
    imageAssets,
    summary,
    messages
  };
}

function createMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [{ role: 'about' }, { type: 'separator' }, { role: 'quit' }]
          }
        ]
      : []),
    {
      label: '文件',
      submenu: [
        { label: '新建', accelerator: 'CmdOrCtrl+N', click: () => emitMenuAction('new-file') },
        { label: '打开...', accelerator: 'CmdOrCtrl+O', click: () => emitMenuAction('open-file') },
        { label: '导入 Word(.docx)...', click: () => emitMenuAction('import-docx') },
        { label: '保存', accelerator: 'CmdOrCtrl+S', click: () => emitMenuAction('save-file') },
        { label: '另存为...', accelerator: 'CmdOrCtrl+Shift+S', click: () => emitMenuAction('save-file-as') },
        { label: '导出 HTML...', accelerator: 'CmdOrCtrl+E', click: () => emitMenuAction('export-html') },
        { label: '导出 PDF...', accelerator: 'CmdOrCtrl+Shift+E', click: () => emitMenuAction('export-pdf') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' }
      ]
    },
    {
      label: '编辑',
      submenu: [
        { role: 'undo', label: '撤销', accelerator: 'CmdOrCtrl+Z' },
        { role: 'redo', label: '重做', accelerator: 'CmdOrCtrl+Y' },
        { type: 'separator' },
        { role: 'cut', label: '剪切', accelerator: 'CmdOrCtrl+X' },
        { role: 'copy', label: '复制', accelerator: 'CmdOrCtrl+C' },
        { role: 'paste', label: '粘贴', accelerator: 'CmdOrCtrl+V' },
        { role: 'selectAll', label: '全选', accelerator: 'CmdOrCtrl+A' }
      ]
    },
    {
      label: '插入',
      submenu: [{ label: '引用图片', accelerator: 'CmdOrCtrl+Shift+I', click: () => emitMenuAction('reference-image') }]
    },
    {
      label: 'AI',
      submenu: [{ label: '打开 AI 对话面板', accelerator: 'CmdOrCtrl+Shift+A', click: () => emitMenuAction('open-ai-chat') }]
    },
    {
      label: '主题',
      submenu: [
        { label: '米白', click: () => emitMenuAction('theme-paper') },
        { label: '雾蓝', click: () => emitMenuAction('theme-mist') },
        { label: '灰砚', click: () => emitMenuAction('theme-slate') },
        { label: '石墨', click: () => emitMenuAction('theme-graphite') },
        { type: 'separator' },
        { label: '终端绿', click: () => emitMenuAction('theme-terminal') },
        { label: '夜码蓝', click: () => emitMenuAction('theme-nightcode') },
        { type: 'separator' },
        { label: '课堂蓝', click: () => emitMenuAction('theme-campus') },
        { label: '青春橙', click: () => emitMenuAction('theme-youth') }
      ]
    },
    {
      label: '视图',
      submenu: [
        { label: '统一工作区', accelerator: 'Alt+1', click: () => emitMenuAction('view-preview') },
        { type: 'separator' },
        { role: 'reload', label: '重新加载', accelerator: 'F5' },
        { role: 'toggledevtools', label: '开发者工具', accelerator: 'F12' }
      ]
    }
  ];

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  allowWindowClose = false;
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1080,
    minHeight: 720,
    show: false,
    backgroundColor: '#efe7d8',
    title: '墨笺 Markdown 编辑器',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow?.show();
    mainWindow?.focus();
  });

  mainWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription) => {
    console.error('[main] failed to load window', errorCode, errorDescription);
  });

  mainWindow.webContents.on('render-process-gone', (_event, details) => {
    console.error('[main] render process gone', details);
    void emergencySaveDocument('render-process-gone');
  });

  const devServerUrl = process.env.VITE_DEV_SERVER_URL;
  console.log('[main] createWindow', { devServerUrl: devServerUrl ?? null });
  if (devServerUrl) {
    mainWindow.loadURL(devServerUrl);
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  mainWindow.on('close', async (event) => {
    if (allowWindowClose || !latestDocumentState.isDirty || isHandlingClose) {
      return;
    }

    event.preventDefault();
    isHandlingClose = true;

    try {
      const response = await dialog.showMessageBox(mainWindow, {
        type: 'question',
        buttons: ['保存', '不保存', '取消'],
        defaultId: 0,
        cancelId: 2,
        title: '是否保存当前文档？',
        message: `关闭前是否保存“${latestDocumentState.displayName || '未命名文档'}”？`,
        detail: latestDocumentState.filePath
          ? '当前文档有未保存修改。'
          : '当前文档还没有保存路径，选择“保存”后会让你选择保存位置。'
      });

      if (response.response === 2) {
        return;
      }

      if (response.response === 0) {
        const saveResult = await persistMarkdownContent(mainWindow, {
          filePath: latestDocumentState.filePath,
          content: latestDocumentState.content,
          forceDialog: !latestDocumentState.filePath,
          defaultPath: latestDocumentState.displayName || 'untitled.md'
        });

        if (saveResult.canceled) {
          return;
        }

        latestDocumentState.filePath = saveResult.filePath;
        latestDocumentState.displayName = path.basename(saveResult.filePath);
        latestDocumentState.isDirty = false;
      }

      allowWindowClose = true;
      mainWindow.close();
    } finally {
      isHandlingClose = false;
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

ipcMain.on('document-state:sync', (_event, payload) => {
  latestDocumentState = {
    filePath: payload?.filePath ?? null,
    displayName: payload?.displayName || '未命名.md',
    content: payload?.content ?? '',
    isDirty: Boolean(payload?.isDirty)
  };
});

ipcMain.handle('dialog:openMarkdownFile', async () => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const { canceled, filePaths } = await dialog.showOpenDialog(targetWindow, {
    title: '打开 Markdown 文件',
    properties: ['openFile'],
    filters: [
      { name: 'Markdown', extensions: ['md', 'markdown', 'mdown', 'txt'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (canceled || filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = filePaths[0];
  const content = await fs.readFile(filePath, 'utf8');
  return { canceled: false, filePath, content };
});

ipcMain.handle('dialog:openDocxFile', async () => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const { canceled, filePaths } = await dialog.showOpenDialog(targetWindow, {
    title: '导入 Word 文档',
    properties: ['openFile'],
    filters: [
      { name: 'Word 文档', extensions: ['docx', 'docm'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (canceled || filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = filePaths[0];
  const preview = await convertDocxFileToPreview(filePath);
  return { canceled: false, ...preview };
});

ipcMain.handle('dialog:saveMarkdownFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  return persistMarkdownContent(targetWindow, payload);
});

ipcMain.handle('dialog:saveImportedMarkdownFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showSaveDialog(targetWindow, {
    title: '导入为 Markdown 文档',
    defaultPath: payload?.defaultPath || 'imported.md',
    filters: [{ name: 'Markdown', extensions: ['md'] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const content = await materializeImportedMarkdown(
    payload?.markdown ?? '',
    payload?.imageAssets ?? [],
    result.filePath
  );

  await fs.writeFile(result.filePath, content, 'utf8');
  return { canceled: false, filePath: result.filePath, content };
});

ipcMain.handle('dialog:materializeImportedMarkdown', async (_event, payload) => {
  if (!payload?.documentPath) {
    throw new Error('缺少目标 Markdown 路径');
  }

  const content = await materializeImportedMarkdown(
    payload?.markdown ?? '',
    payload?.imageAssets ?? [],
    payload.documentPath
  );

  return { canceled: false, content };
});

ipcMain.handle('dialog:readMarkdownFile', async (_event, filePath) => {
  const content = await fs.readFile(filePath, 'utf8');
  return { filePath, content };
});

ipcMain.handle('dialog:exportHtmlFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showSaveDialog(targetWindow, {
    title: '导出 HTML 文件',
    defaultPath: payload?.defaultPath || 'untitled.html',
    filters: [{ name: 'HTML', extensions: ['html'] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  await fs.writeFile(result.filePath, payload?.html ?? '', 'utf8');
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle('dialog:exportPdfFile', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showSaveDialog(targetWindow, {
    title: '导出 PDF 文件',
    defaultPath: payload?.defaultPath || 'untitled.pdf',
    filters: [{ name: 'PDF', extensions: ['pdf'] }]
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  const pdfData = await targetWindow.webContents.printToPDF({
    printBackground: true,
    preferCSSPageSize: true
  });
  await fs.writeFile(result.filePath, pdfData);
  return { canceled: false, filePath: result.filePath };
});

ipcMain.handle('dialog:saveImageAsset', async (_event, payload) => {
  const originalName = payload?.originalName || `image-${Date.now()}.png`;
  const safeName = originalName.replace(/[<>:"/\\|?*\x00-\x1f]/g, '-');
  const parsed = path.parse(safeName);
  const fileName = parsed.ext ? safeName : `${safeName}.png`;

  let targetDir = '';
  if (payload?.documentPath) {
    targetDir = path.join(path.dirname(payload.documentPath), 'assets');
  } else {
    const chosenDir = await dialog.showOpenDialog(BrowserWindow.getFocusedWindow() ?? mainWindow, {
      title: '选择图片保存目录',
      properties: ['openDirectory', 'createDirectory']
    });

    if (chosenDir.canceled || chosenDir.filePaths.length === 0) {
      return { canceled: true };
    }

    targetDir = chosenDir.filePaths[0];
  }

  if (!fsSync.existsSync(targetDir)) {
    await fs.mkdir(targetDir, { recursive: true });
  }

  const filePath = path.join(targetDir, fileName);
  const base64Data = String(payload?.dataUrl || '').replace(/^data:image\/[a-zA-Z0-9+.-]+;base64,/, '');
  const buffer = Buffer.from(base64Data, 'base64');
  await fs.writeFile(filePath, buffer);

  return { canceled: false, filePath };
});

ipcMain.handle('dialog:pickImageReference', async (_event, payload) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  const result = await dialog.showOpenDialog(targetWindow, {
    title: '选择要引用的图片',
    properties: ['openFile'],
    filters: [
      { name: '图片', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = result.filePaths[0];
  let markdownPath = filePath.replace(/\\/g, '/');

  if (payload?.documentPath) {
    const relativePath = path.relative(path.dirname(payload.documentPath), filePath).replace(/\\/g, '/');
    markdownPath = relativePath.startsWith('.') ? relativePath : `./${relativePath}`;
  } else {
    markdownPath = `file:///${markdownPath.replace(/^\/+/, '')}`;
  }

  return {
    canceled: false,
    filePath,
    markdownPath,
    displayName: path.basename(filePath)
  };
});

ipcMain.handle('config:readAiConfigDocument', async () => {
  return readAiConfigDocument();
});

ipcMain.handle('config:writeAiConfigDocument', async (_event, payload) => {
  return writeAiConfigDocument(payload);
});

ipcMain.handle('shell:openExternalLink', async (_event, url) => {
  if (typeof url !== 'string' || !url.trim()) {
    throw new Error('无效链接');
  }

  await shell.openExternal(url);
  return { success: true };
});

app.whenReady().then(() => {
  createMenu();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('session-end', () => {
  void emergencySaveDocument('session-end');
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
