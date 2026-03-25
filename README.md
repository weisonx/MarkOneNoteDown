# MarkOneNoteDown

Windows (WinUI 3) 宸ュ叿锛氬皢 OneNote 绗旇瀵煎嚭骞惰浆鎹负 Markdown锛堝惈鍥剧墖涓庨檮浠讹級銆?
## 鐩爣涓庤寖鍥?
- 鏀寔閫夋嫨绗旇鏈?鍒嗗尯/椤甸潰骞跺鍑轰负 Markdown銆?- 淇濇寔鏍囬灞傜骇銆佸垪琛ㄣ€佽〃鏍笺€佷唬鐮佸潡銆佷换鍔℃竻鍗曠瓑缁撴瀯銆?- 瀵煎嚭鍥剧墖涓庨檮浠跺苟鍦?Markdown 涓敤鐩稿璺緞寮曠敤銆?- 鎻愪緵鍙鍖栬繘搴︿笌鏃ュ織闈㈡澘锛屾敮鎸佹壒閲忓鍑恒€?
## 鎶€鏈€夊瀷

- 璇█涓庤繍琛屾椂锛欳# + .NET 8
- UI锛歐inUI 3
- OneNote 璇诲彇锛歄neNote Interop锛圕OM锛?- Markdown 鐢熸垚锛歁arkdig
- 鏃ュ織锛歋erilog锛堟枃浠?+ UI 鎺у埗鍙帮級
- 閰嶇疆锛歛ppsettings.json + 鐢ㄦ埛璁剧疆
- 鎵撳寘锛歁SIX 鎴?Inno Setup

## 鏋舵瀯涓庢ā鍧?
- UI
  - 绗旇鏈?鍒嗗尯/椤甸潰鏍戦€夋嫨
  - 杈撳嚭鐩綍涓庡鍑洪€夐」閰嶇疆
  - 杩涘害銆佹棩蹇椾笌閿欒鎻愮ず
- OneNoteBridge
  - 灏佽 OneNote COM API
  - 鎷夊彇椤甸潰 XML/HTML 鍐呭
- Parser
  - OneNote XML/HTML 杞崲涓轰腑闂寸粨鏋勶紙鍧楃骇 AST锛?  - 璇嗗埆鏍囬銆佸垪琛ㄣ€佽〃鏍笺€佷唬鐮佸潡銆佷换鍔￠」
- AssetExporter
  - 鍥剧墖/闄勪欢瀵煎嚭鍒?`_assets`
  - 鐢熸垚 Markdown 鐩稿璺緞
- MarkdownRenderer
  - AST 杞?Markdown 鏂囨湰
- ExportPipeline
  - 鎵归噺瀵煎嚭銆侀噸璇曚笌杩涘害鎶ュ憡
- Storage
  - 杈撳嚭鐩綍缁撴瀯涓?OneNote 灞傜骇瀵归綈

## 渚濊禆涓庡墠缃潯浠?
- Windows 10/11
- 宸插畨瑁?OneNote 妗岄潰鐗堬紙鐢ㄤ簬 COM 鎺ュ彛锛?- .NET 8 SDK
- Windows App SDK锛堢敤浜?WinUI 3锛?
### 楠岃瘉瀹夎

鍦?PowerShell 閲屾墽琛岋細  
```powershell
dotnet --info
```
濡傛灉鑳芥甯歌緭鍑?.NET 鐗堟湰淇℃伅锛岃鏄?.NET SDK 鍙敤銆?
鍐嶆墽琛岋細  
```powershell
Get-AppxPackage Microsoft.WindowsAppRuntime*
```
濡傛灉鑳界湅鍒?Microsoft.WindowsAppRuntime.* 鐨勬潯鐩紝灏辫鏄庤繍琛屾椂宸插畨瑁呮垚鍔熴€?
## 鏋勫缓涓庤繍琛岋紙寮€鍙戞満锛?
浣跨敤鍛戒护琛岋細

```powershell
dotnet restore
dotnet build MarkOneNoteDown.sln
```

浣跨敤 Visual Studio锛?
- 鎵撳紑 `MarkOneNoteDown.sln`
- 閫夋嫨 `MarkOneNoteDown.App` 涓哄惎鍔ㄩ」鐩?- 鐩存帴杩愯锛團5锛?
## 鏋勫缓鑴氭湰

椤圭洰鑷甫鑴氭湰 `scripts/build_run_clean.ps1`锛屾敮鎸佹竻鐞嗐€佹瀯寤恒€佽繍琛岋細

```powershell
# 娓呯悊 + 鏋勫缓 + 杩愯锛堥粯璁わ級
powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1

# 鍙竻鐞?powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action clean

# 鍙瀯寤?powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action build

# 鍙繍琛?powershell -ExecutionPolicy Bypass -File scripts\build_run_clean.ps1 -Action run
```


## 瀵煎嚭璇存槑锛堝綋鍓嶅疄鐜帮級

- 椤甸潰鏀寔澶氶€夊鍑猴紙鍦?Pages 鍒楄〃涓寜浣?Ctrl/Shift锛夈€?- 鑻ユ湭閫変腑浠讳綍椤甸潰锛屽垯榛樿瀵煎嚭褰撳墠宸插姞杞界殑鍏ㄩ儴椤甸潰銆?- Markdown 涓哄熀纭€鏂囨湰鎶藉彇鐗堟湰锛堝悗缁細鎵╁睍鏍囬/鍒楄〃/琛ㄦ牸/鍥剧墖鏄犲皠锛夈€?
## OneNote HTML 瀵煎嚭娴佺▼锛堟帹鑽愶級

1. 鍦?OneNote 妗岄潰鐗堥€夋嫨绗旇鏈垨鍒嗗尯銆?2. 浣跨敤 OneNote 鐨勨€滃鍑衡€濅负 HTML锛堢敓鎴愪竴涓寘鍚涓?`*.html` 鐨勬枃浠跺す锛夈€?3. 鍦ㄥ簲鐢ㄤ腑閫夋嫨璇ュ鍑烘枃浠跺す浣滀负 Source銆?4. 閫夋嫨杈撳嚭鐩綍骞跺鍑轰负 Markdown銆?
## PDF 鏀寔锛堟柊澧烇級

褰撳鍑虹洰褰曚腑鍖呭惈 `*.pdf` 鏂囦欢鏃讹紝浼氳繘琛屾枃鏈彁鍙栧苟杞崲涓?Markdown銆?
娉ㄦ剰锛氬綋鍓嶄娇鐢?iText锛圓GPL 鍟嗕笟鎺堟潈瑕佹眰锛夎繘琛?PDF 鏂囨湰鎻愬彇锛岃嫢鐢ㄤ簬闂簮/鍟嗕笟椤圭洰璇风‘淇濈鍚堣鍙瘉銆?
## 閰嶇疆锛堝彲閫夛級

鍙湪 `appsettings.json` 涓厤缃粯璁ゆ簮鐩綍锛?
```json
{
  "SourceFolder": ""
}
```

- `SourceFolder` 闈炵┖鏃跺惎鍔ㄨ嚜鍔ㄥ姞杞借鐩綍鐨?HTML 椤甸潰銆?
## 浣跨敤 VS Code锛堟瀯寤轰笌杩愯锛?
鍑嗗宸ヤ綔锛?
- 瀹夎 VS Code 鎵╁睍锛欳# Dev Kit锛堝寘鍚?C#銆?NET 璋冭瘯鏀寔锛?- 纭宸插畨瑁?.NET 8 SDK 涓?Windows App SDK 杩愯鏃?
鏋勫缓锛?
```powershell
dotnet restore
dotnet build MarkOneNoteDown.sln
```

杩愯锛堝惎鍔?WinUI 3 搴旂敤锛夛細

```powershell
dotnet run --project src\MarkOneNoteDown.App\MarkOneNoteDown.App.csproj
```

濡傛灉杩愯鏃舵姤鏉冮檺鎴栭儴缃茬浉鍏抽敊璇紝璇风‘璁わ細

- Windows 寮€鍙戣€呮ā寮忓凡寮€鍚?- Windows App Runtime 宸叉纭畨瑁?
## 璁″垝涓殑鐩綍缁撴瀯锛堝缓璁級

```
src/
  MarkOneNoteDown.App/          WinUI 3 鍏ュ彛涓?UI
  MarkOneNoteDown.Core/         瑙ｆ瀽涓庢覆鏌撴牳蹇冮€昏緫
  MarkOneNoteDown.OneNote/      COM 浜掓搷浣滃皝瑁?  MarkOneNoteDown.Export/       瀵煎嚭娴佺▼涓庝换鍔¤皟搴?  MarkOneNoteDown.Tests/        鍗曞厓涓庨泦鎴愭祴璇?```

## 瀹炵幇閲岀▼纰?
1. 寤虹珛 WinUI 3 搴旂敤楠ㄦ灦涓庡熀纭€瀵艰埅椤甸潰
2. OneNote COM 杩炴帴涓庢爲褰㈢粨鏋勮鍙?3. 鍗曢〉瀵煎嚭涓?Markdown 娓叉煋锛圡VP锛?4. 鎵归噺瀵煎嚭銆侀檮浠舵敮鎸佷笌杩涘害鏄剧ず
5. 绋冲畾鎬т笌鎬ц兘浼樺寲銆佸畨瑁呭寘鍙戝竷

## 澶囨敞

OneNote 鍘熺敓 `.one`/`.onetoc2` 鏂囦欢涓嶅彲鐩存帴瑙ｆ瀽锛屾帹鑽愰€氳繃 OneNote COM API 璇诲彇鍐呭骞惰浆鎹负 Markdown銆?
## EXE 安装包（Inno Setup）

如果需要一个可双击安装的 EXE，可以使用 Inno Setup 作为引导安装器，它会：
1. 导入签名证书到本机信任库
2. 安装 MSIX 包

准备：
- 安装 Inno Setup（确保 `ISCC.exe` 在 PATH 中）

打包：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\package_exe.ps1
```

输出位置：
- `artifacts\installer\MarkOneNoteDown-Setup.exe`

注意：
- 安装过程需要管理员权限（用于导入证书）。

## EXE 鎵撳寘锛堜笉浣跨敤 MSIX锛?
宸叉敼涓轰紶缁?EXE 鍙戝竷娴佺▼锛屼笉鍐嶇敓鎴?MSIX 瀹夎鍖呫€傝浣跨敤浠ヤ笅鑴氭湰鐢熸垚鍙繍琛岀殑 EXE 鐩綍锛?
```powershell
powershell -ExecutionPolicy Bypass -File scripts\publish_exe.ps1
```

杈撳嚭浣嶇疆锛?- `artifacts\exe\`
- 杩愯 `artifacts\exe\MarkOneNoteDown.App.exe`

鍙€夊弬鏁帮細
- `-Configuration Release|Debug`
- `-Runtime win-x64`
- `-SelfContained:$true|$false`
- `-OutputDir "artifacts\exe"`


## 单文件安装包（Inno Setup）

如果需要单文件安装包（EXE 安装器），请先安装 Inno Setup 并确保 `ISCC.exe` 在 PATH 中。

生成安装包：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\package_installer.ps1
```

输出位置：
- `artifacts\installer\MarkOneNoteDown-Setup.exe`

说明：
- 脚本会先执行 `publish_exe.ps1` 生成发布目录，再打包成安装器。
