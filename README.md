# 墨笺 Markdown

Mojian Markdown 是一款面向 Windows 的本地优先 Markdown 桌面编辑器，适合写技术文档、教程、方案说明和日常笔记。项目基于 Electron、React、Vite 和 TypeScript，强调本地文件可控、中文友好、即时预览，以及可选的 AI 辅助写作能力。

[GitHub Releases](https://github.com/codeRunning557/mojian-markdown-editor/releases)

## 当前可用功能

- 打开、创建、保存、另存为本地 Markdown 文件
- 富文本编辑区、实时预览与标题大纲联动
- 支持图片引用、拖拽插入、粘贴插入
- 代码块语言标签与一键复制
- 导出 HTML、导出 PDF
- 导入 Word `.docx`，预览转换后的 Markdown，并在保存时写入图片资源
- AI 对话面板，支持改写、扩写、精简、续写
- 文档或选区翻译预览
- 内置多套 AI 服务预设，也支持自定义 OpenAI-compatible 接口
- 提供 AI 请求日志、主题切换、最近文件和异常退出自动恢复

## 安装与使用

普通用户建议直接从 Releases 下载打包版本，而不是从源码启动。当前发布页对应的主要交付物为：

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

如果你只想使用软件：

1. 从 Releases 下载上述任一 Windows 发布物。
2. 解压 `墨笺 Markdown.zip`，或直接运行 `墨笺 Markdown.exe`。
3. 首次运行若被系统提示校验或权限确认，按系统提示继续即可。

## 从源码运行

建议使用较新的 Node.js LTS 版本。

```bash
npm install
npm run dev
```

`npm run dev` 会启动 Vite 开发服务器并拉起 Electron。

如需先构建再启动桌面端：

```bash
npm start
```

## 打包方式

项目当前可用的 Windows 打包命令如下：

```bash
npm run pack:win
npm run dist:win
npm run dist:zip
npm run dist:portable
```

说明：

- `npm run pack:win` 生成目录版应用，输出到 `release/win-unpacked`
- `npm run dist:win` 生成 Windows 安装包
- `npm run dist:zip` 生成 zip 发布包
- `npm run dist:portable` 生成便携版可执行文件

## 技术栈

- Electron
- React 19
- Vite
- TypeScript
- `marked`
- `DOMPurify`
- `mammoth`
- `turndown`

## 仓库结构

```text
electron/   Electron 主进程与预加载脚本
scripts/    开发启动与桌面端辅助脚本
src/        React 前端源码
docs/       补充文档
release/    本地打包输出目录
```

## 说明

- 应用目前以 Windows 桌面端为主
- Markdown 文件始终保存在本地，AI 配置也会优先保存在用户私有位置
- 仓库首页说明以当前代码实现为准，发布页以最新 Release 为准
