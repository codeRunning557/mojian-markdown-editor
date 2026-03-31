# 墨笺 Markdown 编辑器

墨笺是一款本地优先的 Markdown 桌面编辑器，面向技术文档、教程文档、方案文档和日常写作场景。

它基于 Electron、React、Vite 和 TypeScript 构建，强调极简界面、Markdown 文件可控、中文友好，以及 AI 辅助写作能力。

## 当前功能

- 打开、保存、另存为本地 Markdown 文件
- 富文本编辑工作区，内容最终保存为 `.md`
- 左侧标题大纲与正文跳转
- 图片引用、图片粘贴、图片拖拽插入
- 代码块语言标签与复制按钮
- 导出 HTML
- 导出 PDF
- AI 美化预览后应用
- 翻译预览后应用
- 右键菜单插入标题、代码块、图片
- 多套极简主题切换

## 技术栈

- Electron
- React 19
- Vite
- TypeScript

## 本地开发

```bash
npm install
npm run dev
```

开发模式会启动 Vite 和 Electron。

## 本地运行

```bash
npm start
```

`npm start` 会先构建，再启动桌面应用。

## 构建目录版应用

```bash
npm run pack:win
```

构建结果输出到 `release/win-unpacked`。

## 项目结构

```text
electron/   Electron 主进程与预加载脚本
scripts/    启动与构建辅助脚本
src/        React 前端源码
```

## 适合的使用场景

- 技术部署文档
- 运维排障文档
- API 或方案说明文档
- Markdown 日常写作

## 说明

仓库默认只保留运行和开发所需代码与配置，不包含内部产品规划文档。
