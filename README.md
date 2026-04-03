# 墨笺 Markdown

本地优先的 Windows Markdown 桌面编辑器，基于 Electron、React、Vite 和 TypeScript 构建。它面向长文写作、技术文档整理和本地资料编辑，提供实时预览、Word 导入、HTML/PDF 导出，以及可选的 AI 辅助能力。

[下载最新版本](https://github.com/codeRunning557/mojian-markdown-editor/releases)

## 当前功能

- 本地文件工作流：新建、打开、保存、另存为 Markdown 文件
- 实时渲染编辑体验：统一工作区写作视图、标题大纲联动、当前章节定位
- 富文本式编辑：列表、引用、代码块、内联代码、右键 Markdown 语法插入
- 图片处理：引用本地图片、拖拽插入、粘贴插入，并按文档相对路径组织图片资源
- 代码内容增强：代码块语言标签、一键复制、编辑区与预览区统一代码样式
- 文档导出：导出 HTML、导出 PDF
- Word 导入：导入 `.docx` / `.docm`，先预览转换结果，再导入为新文档或插入当前文档
- AI 辅助：内置离线文档助手，以及可配置的 OpenAI-compatible 接口
- AI 工作流：改写、扩写、精简、续写、翻译、美化预览、连接测试和请求日志
- 写作辅助：最近文件记录、撤销/重做、多主题切换、异常退出自动恢复

## 安装与使用

项目当前面向 Windows 发布，公开发布物为：

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

使用方式：

1. 在 [GitHub Releases](https://github.com/codeRunning557/mojian-markdown-editor/releases) 下载任一发布物。
2. 直接运行 `墨笺 Markdown.exe`，或解压 `墨笺 Markdown.zip` 后启动应用。
3. 如果 Windows 首次启动时出现安全提示，确认后继续即可。

## 本地开发

建议使用较新的 Node.js LTS 版本。

```bash
npm install
npm run dev
```

如果需要先构建再启动 Electron：

```bash
npm start
```

## 打包与发布产物

可用打包命令如下：

```bash
npm run pack:win
npm run dist:win
npm run dist:zip
npm run dist:portable
```

对应说明：

- `npm run pack:win`：生成目录版应用，输出到 `release/win-unpacked`
- `npm run dist:win`：生成 Windows 安装包
- `npm run dist:zip`：生成 Windows 压缩包
- `npm run dist:portable`：生成便携版 `.exe`

当前发布流程会从仓库根目录上传以下两个文件作为 GitHub Release 资产：

- `墨笺 Markdown.exe`
- `墨笺 Markdown.zip`

## 技术栈

- Electron
- React 19
- Vite
- TypeScript
- `marked`
- `DOMPurify`
- `mammoth`
- `turndown`

## 说明

- 应用以本地文件系统为中心，不依赖云端文档存储。
- AI 模型档案模板与真实密钥分离保存，真实配置保存在当前用户私有目录。
- 仓库内的发布目录和根目录发布物主要服务于 Windows 桌面分发流程。