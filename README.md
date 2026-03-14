# MarkOneNoteDown

Windows (WinUI 3) 工具：将 OneNote 笔记导出并转换为 Markdown（含图片与附件）。

## 目标与范围

- 支持选择笔记本/分区/页面并导出为 Markdown。
- 保持标题层级、列表、表格、代码块、任务清单等结构。
- 导出图片与附件并在 Markdown 中用相对路径引用。
- 提供可视化进度与日志面板，支持批量导出。

## 技术选型

- 语言与运行时：C# + .NET 8
- UI：WinUI 3
- OneNote 读取：OneNote Interop（COM）
- Markdown 生成：Markdig
- 日志：Serilog（文件 + UI 控制台）
- 配置：appsettings.json + 用户设置
- 打包：MSIX 或 Inno Setup

## 架构与模块

- UI
  - 笔记本/分区/页面树选择
  - 输出目录与导出选项配置
  - 进度、日志与错误提示
- OneNoteBridge
  - 封装 OneNote COM API
  - 拉取页面 XML/HTML 内容
- Parser
  - OneNote XML/HTML 转换为中间结构（块级 AST）
  - 识别标题、列表、表格、代码块、任务项
- AssetExporter
  - 图片/附件导出到 `_assets`
  - 生成 Markdown 相对路径
- MarkdownRenderer
  - AST 转 Markdown 文本
- ExportPipeline
  - 批量导出、重试与进度报告
- Storage
  - 输出目录结构与 OneNote 层级对齐

## 依赖与前置条件

- Windows 10/11
- 已安装 OneNote 桌面版（用于 COM 接口）
- .NET 8 SDK
- Windows App SDK（用于 WinUI 3）

## 验证安装

在 PowerShell 里执行：  
```powershell
dotnet --info
```
如果能正常输出 .NET 版本信息，说明 .NET SDK 可用。

再执行：  
```powershell
Get-AppxPackage Microsoft.WindowsAppRuntime*
```
如果能看到 Microsoft.WindowsAppRuntime.* 的条目，就说明运行时已安装成功。

## 计划中的目录结构（建议）

```
src/
  MarkOneNoteDown.App/          WinUI 3 入口与 UI
  MarkOneNoteDown.Core/         解析与渲染核心逻辑
  MarkOneNoteDown.OneNote/      COM 互操作封装
  MarkOneNoteDown.Export/       导出流程与任务调度
  MarkOneNoteDown.Tests/        单元与集成测试
```

## 实现里程碑

1. 建立 WinUI 3 应用骨架与基础导航页面
2. OneNote COM 连接与树形结构读取
3. 单页导出与 Markdown 渲染（MVP）
4. 批量导出、附件支持与进度显示
5. 稳定性与性能优化、安装包发布

## 备注

OneNote 原生 `.one`/`.onetoc2` 文件不可直接解析，推荐通过 OneNote COM API 读取内容并转换为 Markdown。
