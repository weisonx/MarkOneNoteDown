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

### 验证安装

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

## 构建与运行（开发机）

使用命令行：

```powershell
dotnet restore
dotnet build MarkOneNoteDown.sln
```

使用 Visual Studio：

- 打开 `MarkOneNoteDown.sln`
- 选择 `MarkOneNoteDown.App` 为启动项目
- 直接运行（F5）

## 构建脚本

项目自带脚本 `scripts/build_run_clean.ps1`，支持清理、构建、运行：

```powershell
# 清理 + 构建 + 运行（默认）
powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1

# 只清理
powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action clean

# 只构建
powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action build

# 只运行
powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action run
```

## 导出说明（当前实现）

- 页面支持多选导出（在 Pages 列表中按住 Ctrl/Shift）。
- 若未选中任何页面，则默认导出当前已加载的全部页面。
- Markdown 为基础文本抽取版本（后续会扩展标题/列表/表格/图片映射）。

## OneNote HTML 导出流程（推荐）

1. 在 OneNote 桌面版选择笔记本或分区。
2. 使用 OneNote 的“导出”为 HTML（生成一个包含多个 `*.html` 的文件夹）。
3. 在应用中选择该导出文件夹作为 Source。
4. 选择输出目录并导出为 Markdown。

## PDF 支持（新增）

当导出目录中包含 `*.pdf` 文件时，会进行文本提取并转换为 Markdown。

注意：当前使用 iText（AGPL 商业授权要求）进行 PDF 文本提取，若用于闭源/商业项目请确保符合许可证。

## 配置（可选）

可在 `appsettings.json` 中配置默认源目录：

```json
{
  "SourceFolder": ""
}
```

- `SourceFolder` 非空时启动自动加载该目录的 HTML 页面。

## 使用 VS Code（构建与运行）

准备工作：

- 安装 VS Code 扩展：C# Dev Kit（包含 C#、.NET 调试支持）
- 确认已安装 .NET 8 SDK 与 Windows App SDK 运行时

构建：

```powershell
dotnet restore
dotnet build MarkOneNoteDown.sln
```

运行（启动 WinUI 3 应用）：

```powershell
dotnet run --project src\MarkOneNoteDown.App\MarkOneNoteDown.App.csproj
```

如果运行时报权限或部署相关错误，请确认：

- Windows 开发者模式已开启
- Windows App Runtime 已正确安装

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
