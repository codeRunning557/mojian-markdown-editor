# 墨笺 Markdown / 墨箋 Markdown / Mojian Markdown

[Download Latest Release](https://github.com/codeRunning557/mojian-markdown-editor/releases)

---

## 简体中文

墨笺 Markdown 是一款面向 Windows 的本地优先 Markdown 桌面编辑器，基于 Electron、React、Vite 和 TypeScript 构建。它专注于长文写作、技术文档整理、Word 导入转换，以及可选的 AI 辅助写作能力。

### 当前功能

- 本地 Markdown 文件的新建、打开、保存、另存为
- 富文本编辑工作区与单栏阅读式写作体验
- 左侧大纲导航与当前标题联动
- 右侧 AI 会话面板与多模型配置
- 图片引用、拖拽插入、粘贴插入
- 代码块语言标签与一键复制
- 导出 HTML / PDF
- 导入 Word `.docx`，预览转换结果后再导入
- AI 改写、扩写、精简、续写、翻译预览
- 多主题切换、右键 Markdown 语法插入、撤销 / 重做

### 下载与使用

GitHub Release 中提供两个对外发布物：

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

使用方式：

1. 从 Releases 页面下载任一 Windows 发布物
2. 直接运行 `墨笺 Markdown.exe`，或解压 `墨笺 Markdown.zip` 后启动应用
3. 如果 Windows 首次启动提示安全确认，允许继续即可

### 本地开发

建议使用较新的 Node.js LTS 版本。

```bash
npm install
npm run dev
```

如果需要先构建再启动桌面应用：

```bash
npm start
```

### 打包命令

```bash
npm run pack:win
npm run dist:win
npm run dist:zip
npm run dist:portable
```

输出说明：

- `npm run pack:win`：生成目录版，输出到 `release/win-unpacked`
- `npm run dist:win`：生成 Windows 安装版
- `npm run dist:zip`：生成 Windows 压缩包
- `npm run dist:portable`：生成便携版 `.exe`

### 技术栈

- Electron
- React 19
- Vite
- TypeScript
- `marked`
- `DOMPurify`
- `mammoth`
- `turndown`

### 说明

- 当前版本以 Windows 桌面端为主
- Markdown 文档始终保留在本地文件系统
- AI 配置模板与真实密钥分离保存，真实密钥不会随公开发布包一起分发

---

## 繁體中文

墨箋 Markdown 是一款面向 Windows 的本地優先 Markdown 桌面編輯器，基於 Electron、React、Vite 與 TypeScript 建構。它專注於長文寫作、技術文件整理、Word 匯入轉換，以及可選的 AI 輔助寫作能力。

### 目前功能

- 本地 Markdown 檔案的新建、開啟、儲存、另存新檔
- 富文本編輯工作區與單欄閱讀式寫作體驗
- 左側大綱導航與當前標題聯動
- 右側 AI 對話面板與多模型設定
- 圖片引用、拖曳插入、貼上插入
- 程式碼區塊語言標籤與一鍵複製
- 匯出 HTML / PDF
- 匯入 Word `.docx`，預覽轉換結果後再匯入
- AI 改寫、擴寫、精簡、續寫、翻譯預覽
- 多主題切換、右鍵 Markdown 語法插入、復原 / 重做

### 下載與使用

GitHub Release 提供兩個對外發佈物：

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

使用方式：

1. 從 Releases 頁面下載任一 Windows 發佈物
2. 直接執行 `墨笺 Markdown.exe`，或解壓 `墨笺 Markdown.zip` 後啟動應用
3. 若 Windows 首次啟動顯示安全提示，允許繼續即可

### 本地開發

建議使用較新的 Node.js LTS 版本。

```bash
npm install
npm run dev
```

若需要先建構再啟動桌面應用：

```bash
npm start
```

### 打包命令

```bash
npm run pack:win
npm run dist:win
npm run dist:zip
npm run dist:portable
```

輸出說明：

- `npm run pack:win`：產生目錄版，輸出到 `release/win-unpacked`
- `npm run dist:win`：產生 Windows 安裝版
- `npm run dist:zip`：產生 Windows 壓縮包
- `npm run dist:portable`：產生便攜版 `.exe`

### 技術棧

- Electron
- React 19
- Vite
- TypeScript
- `marked`
- `DOMPurify`
- `mammoth`
- `turndown`

### 說明

- 目前版本以 Windows 桌面端為主
- Markdown 文件始終保留在本地檔案系統
- AI 設定模板與真實金鑰分離保存，真實金鑰不會隨公開發佈包一同分發

---

## English

Mojian Markdown is a local-first Markdown desktop editor for Windows, built with Electron, React, Vite, and TypeScript. It focuses on long-form writing, technical documentation, Word import/conversion, and optional AI-assisted writing.

### Current Features

- Create, open, save, and save-as local Markdown files
- Rich-text editing workspace with a single-column reading-oriented writing mode
- Left-side outline navigation with current-heading linkage
- Right-side AI chat panel with multi-model configuration
- Image reference, drag-and-drop insertion, and paste insertion
- Code block language labels with one-click copy
- Export to HTML / PDF
- Import Word `.docx`, preview the converted Markdown, then import it
- AI rewrite, expand, shorten, continue, and translation preview
- Multiple themes, right-click Markdown syntax insertion, undo / redo

### Download and Use

GitHub Releases provides two public Windows artifacts:

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

How to use:

1. Download either Windows artifact from the Releases page
2. Run `墨笺 Markdown.exe`, or extract `墨笺 Markdown.zip` and start the app
3. If Windows shows a first-launch security prompt, allow it to continue

### Local Development

Use a recent Node.js LTS release.

```bash
npm install
npm run dev
```

To build first and then start the desktop app:

```bash
npm start
```

### Packaging Commands

```bash
npm run pack:win
npm run dist:win
npm run dist:zip
npm run dist:portable
```

Outputs:

- `npm run pack:win`: unpacked app directory in `release/win-unpacked`
- `npm run dist:win`: Windows installer
- `npm run dist:zip`: Windows zip package
- `npm run dist:portable`: portable `.exe`

### Tech Stack

- Electron
- React 19
- Vite
- TypeScript
- `marked`
- `DOMPurify`
- `mammoth`
- `turndown`

### Notes

- The current release is Windows-first
- Markdown documents remain on the local filesystem
- AI profile templates and real secrets are stored separately, and real keys are not distributed in public packages
