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

## 普通用户使用

如果你不打算参与开发，建议直接下载打包好的 Windows 可执行版本，而不是从源码启动。

### 下载方式

- 前往 GitHub Releases 页面下载最新发布包
- Releases 页面：`https://github.com/codeRunning557/mojian-markdown-editor/releases`
- 优先下载 `win-unpacked` 对应压缩包，或后续提供的便携版 / 安装版

### 启动方式

- 解压发布包
- 双击 `墨笺 Markdown 编辑器.exe`
- 首次运行如果被系统提示，请选择“仍要运行”或加入信任

## 构建目录版应用

```bash
npm run pack:win
```

构建结果输出到 `release/win-unpacked`。

## 构建发布压缩包

```bash
npm run dist:zip
```

适合给不想安装的用户提供备用下载。

## 构建 Windows 安装包

```bash
npm run dist:win
```

推荐把 `dist:win` 产出的 NSIS 安装包上传到 GitHub Releases。对于普通用户，它通常比目录版或 portable 更小，也更容易安装。

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

## 发布建议

如果你要把应用分发给普通用户，建议通过 GitHub Releases 发布，而不是直接让用户下载源码。

推荐流程：

1. 本地执行 `npm run pack:win`
2. 面向普通用户分发时，优先执行 `npm run dist:win`
3. 如需免安装备用包，再执行 `npm run dist:zip`
4. 在 GitHub 仓库创建一个新的 Release
5. 优先上传 NSIS 安装包，zip 可作为备用下载
6. 参考仓库根目录的 `RELEASE_NOTES_TEMPLATE.md` 填写发布说明
