---
name: AHK v2 开发与跨账号无缝接续工作流
description: 创建、修改、重构或调试 AutoHotkey v2 项目时使用。强制先规划并等待授权，维护项目概览以实现跨账号接续，版本隔离保存，并自动运行验证。
---

# AutoHotkey v2 开发与跨账号无缝接续工作流规范

本技能适用于所有 AutoHotkey v2 项目的创建、修改、重构与调试。执行时以本规范为准；未经用户明确授权，不得创建或修改任何代码文件。

## 1. 基础环境与交互规范

- **语言要求**：全程严格使用中文交流。
- **Git 安装路径**：`W:\git`
  - 命令行：`W:\git\cmd\git.exe`
  - 本机已验证版本：`git version 2.56.0.windows.1`
- **AutoHotkey 安装路径**：`C:\AutoHotkey`
  - v2 解释器：`C:\AutoHotkey\v2\AutoHotkey64.exe`
  - 编译器：`C:\AutoHotkey\Compiler\Ahk2Exe.exe`
  - 本机已验证版本：`2.0.28`（升级后以实际版本为准）
- **Python 安装路径**：`C:\Python`
  - 插件：`如果需要安装什么插件，请自行安装`
- **语法规范**：100% 使用 AutoHotkey v2 语法，严禁混用 v1。
- **代码注释**：所有函数、核心分支必须包含详细的中文注释。
- **自动运行**：修改完成后，自动调用 AHK 运行程序供用户即时测试。
- **结束标识**：修改代码的对话输出完毕后，最后一行必须单独提示：`已完成`。非代码修改的常规问答对话严禁输出该提示。

## 2. 权限与执行顺序（两阶段控制）

1. **方案规划先行**：接收到需求后，先输出实现思路与规划。
2. **严格等待授权**：在用户明确回复“可以编写代码”或类似确认指令前，绝对严禁创建或修改任何代码文件。
3. **版本隔离保存**：经许可修改代码时，不得覆盖原程序，必须递增版本号另存为新文件，并在新文件顶部写入本次变更日志（ChangeLog）。

授权判断以肯定指令为准，例如“可以编写代码”“开始写”“按此实现”。用户的沉默、继续提问或修改规划都不算授权。

版本隔离的具体要求：

- 入口脚本命名为 `main_vX.Y.Z.ahk`，版本号递增，旧版本必须保留。
- 新文件顶部必须写入中文 ChangeLog，说明本次改了什么、为什么改。
- 被修改的库文件同样要保留旧版，或在 Git 中提交并打标签；严禁只保留一份被覆盖的副本。
- 版本号、当前文件路径必须同步写入概览文件和变更记录。

## 3. 防中断与跨账号交接规则（核心必须执行）

当前 AI 账号存在调用时间与额度限制，随时可能中断。为确保切换账号后无缝接力，必须严格执行以下接续协议。

### 阶段 A：修改代码前的“前置落盘”（强制阻断门）

- 当用户同意开始修改代码后，**在动任何代码文件之前**，必须先在项目根目录下创建或更新《XX项目整体情况概览.md》。
- 实际文件名用项目名替换 `XX`，例如 `MyProject整体情况概览.md`；不要把书名号写进文件名。
- **绝不能等写完代码再更新**；若预判调用时间或上下文额度紧张，**优先确保该概览文件处于最新、可接续的状态**。
- 概览文件更新完成后，才能开始修改代码。
- 若概览文件已存在但不完整，先补全 10 项内容再动代码，不得凭猜测继续。

### 阶段 B：《XX项目整体情况概览.md》标准结构

概览文件必须具体、准确、可执行，严禁空泛描述。涉及路径、命令、依赖、配置的，必须写明。必须包含以下 10 项：

1. **项目目标与范围**：该工具解决什么问题、核心功能边界。
2. **当前整体进度**：明确标注各模块是已完成、进行中还是待启动。
3. **已完成内容**：本次修改前已稳定运行的功能列表。
4. **未完成/待修改内容**：当前亟需解决的缺陷或待拓展功能。
5. **下一步具体修改点**：按优先级列出具体文件路径、函数/模块名、预期输入输出及修改结果。
6. **关键文件与目录说明**：当前最新版本文件名称、依赖脚本路径。
7. **运行与验证命令**：附带绝对路径的启动与调试命令，例如 `C:\AutoHotkey\v2\AutoHotkey64.exe "W:\project\main_v1.0.2.ahk"`。
8. **已知问题、风险与注意事项**：临时规避方案、潜在的语法或并发坑点。
9. **最近变更记录**：记录上一版本到当前版本的演进差异。
10. **给下一个 AI 账号的接续指令**：告知新账号接手后具体从第几行的哪个函数继续编写。

### 阶段 C：新账号接盘响应流程

当用户更换账号或开启新会话时，新账号必须按以下顺序行动：

1. 主动读取项目根目录下的《XX项目整体情况概览.md》。
2. 首先向用户简要复述：当前项目进展、上个账号停留的节点、准备执行的下一步计划。
3. 等待用户确认或收到指令后，直接从概览文件中的“下一步修改点”继续推进。

## 4. 技术执行参考（获授权后使用）

本节只提供获授权后的实现参考，不替代第 2 节的授权门和第 3 节的前置落盘要求。

### 4.1 项目结构与版本文件

```text
MyProject/
  MyProject整体情况概览.md
  main_v0.1.0.ahk          ; 当前版本入口，每次修改递增版本号
  lib/
    App.ahk                ; 应用名称、版本等常量
    Log.ahk                ; 统一日志
  tests/
    run-tests.ahk          ; 无界面测试入口，失败返回非 0
  assets/
    app.ico
  dist/                    ; 编译输出
  README.md
  CHANGELOG.md
```

每次修改入口脚本时复制为新版本文件，例如 `main_v0.1.1.ahk`，不要覆盖 `main_v0.1.0.ahk`。`lib/` 中被修改的文件也要保留旧版或在 Git 中打标签。

### 4.2 入口文件最低要求

```autohotkey
; ============================================================
; 文件：main_v0.1.0.ahk
; 项目：MyProject
; 版本：0.1.0
; 变更记录（ChangeLog）：
;   0.1.0 - 初始版本：创建入口、日志和全局错误处理
; ============================================================
#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, OutputDebug

#Include lib\App.ahk
#Include lib\Log.ahk

Main()

; 主流程：初始化运行环境并启动程序
Main() {
    SetWorkingDir A_ScriptDir
    SendMode "Input"
    CoordMode "Mouse", "Screen"
    OnError FatalError

    Log("启动 " App.Name " " App.Version)
    Persistent()   ; 已有热键或界面后可删除
}

; 全局错误处理：记录堆栈并向用户显示可读错误
FatalError(err, mode) {
    Log("致命错误 " err.Stack)
    MsgBox err.Message, App.Name " 错误", "Iconx"
    ExitApp 1
}
```

```autohotkey
; lib/App.ahk：全局常量
class App {
    static Name := "MyProject"
    static Version := "0.1.0"
}
```

```autohotkey
; lib/Log.ahk：统一日志，写入用户 AppData，单文件超过 5 MB 时轮转
Log(message) {
    dir := A_AppData "\" App.Name
    file := dir "\app.log"
    DirCreate dir
    if (FileExist(file) && FileGetSize(file) > 5000000)
        FileMove file, file ".1", true
    FileAppend FormatTime(, "yyyy-MM-dd HH:mm:ss") " " message "`n", file, "UTF-8"
}
```

用户数据写入 `A_AppData` 或 `A_Temp`，不要写脚本所在目录；编译后的程序可能位于 `Program Files`，没有写权限。

### 4.3 AHK v2 编码约束

- 每个入口文件首部必须有 `#Requires AutoHotkey v2.0`。
- `:=` 是赋值，`=` 是大小写不敏感比较，`==` 是大小写敏感比较。
- 禁止 v1 命令语法，统一使用 v2 函数和表达式。
- 热键体超过一条语句时必须用花括号；上下文热键用 `#HotIf` 开启，并用空的 `#HotIf` 关闭。
- GUI 对象必须用全局变量或静态字段保持引用；对象被回收后窗口会消失。
- 用 `try/catch/finally` 处理异常，严禁空的 `catch`。
- 路径优先使用 `A_ScriptDir`、`A_AppData`、`A_Temp`，不要依赖当前工作目录。
- 发送字面文本用 `SendText`，发送按键用 `Send`。
- 所有函数和核心分支写中文注释，说明输入、输出和失败行为。

常见 v1 → v2 对照：

| v1                                                       | v2                            |
| -------------------------------------------------------- | ----------------------------- |
| `MsgBox, 文本`                                           | `MsgBox "文本"`               |
| `SetTimer, Label, 1000`                                  | `SetTimer Label, 1000`        |
| `Gui, Add, Text`                                         | `g.AddText()`                 |
| `StringReplace, out, in, a, b`                           | `out := StrReplace(in, a, b)` |
| `IfWinActive, 标题`                                      | `if WinActive("标题")`        |
| `#IfWinActive 标题`                                      | `#HotIf WinActive("标题")`    |
| `Clipboard`                                              | `A_Clipboard`                 |
| `FileRead, out, 文件`                                    | `out := FileRead(文件)`       |
| `Loop, %count%`                                          | `Loop count`                  |
| `Gosub, Label`                                           | 调用函数                      |
| `ByRef` 参数                                             | `&参数`                       |
| `#NoEnv`、`SetBatchLines`、`StringCaseSense`、`AutoTrim` | 已移除，删除或替换            |

### 4.4 调试与验证

```powershell
# 前台运行：适合控制台/短脚本，能直接看到报错
& "C:\AutoHotkey\v2\AutoHotkey64.exe" "W:\project\main_v0.1.0.ahk"

# 后台运行：适合常驻热键/GUI 脚本，避免阻塞当前会话
Start-Process "C:\AutoHotkey\v2\AutoHotkey64.exe" -ArgumentList '"W:\project\main_v0.1.0.ahk"'
```

运行后，把实际结果、退出码和用户需要测试的步骤写入概览文件的第 7、10 项。

常见问题优先检查：

| 现象                     | 优先检查                                               |
| ------------------------ | ------------------------------------------------------ |
| 脚本启动后立即退出       | 没有热键/GUI，且没有调用 `Persistent()`                |
| 热键不触发               | `#HotIf` 条件、重复实例、其他脚本占用、目标窗口权限    |
| `Send` 无效              | 目标窗口焦点、`SendMode`、目标窗口是否以管理员权限运行 |
| 点击位置偏移             | `CoordMode`、屏幕/窗口坐标、DPI 缩放                   |
| GUI 一闪而过             | 没有全局/静态引用 GUI 对象                             |
| 开发环境正常、编译后失败 | 相对路径、`A_ScriptDir`、资源是否打包、文件权限        |

### 4.5 测试

纯逻辑放在 `lib/`，测试入口保持无界面：不弹 `MsgBox`、不依赖热键、不读写真实用户数据。

```autohotkey
; tests/run-tests.ahk：无界面测试入口，失败时返回非 0
#Requires AutoHotkey v2.0
#SingleInstance Force

failures := 0
logFile := A_ScriptDir "\test-results.log"
if FileExist(logFile)
    FileDelete logFile

Check("1+1=2", 1 + 1 = 2, logFile)
Check("字符串替换", StrReplace("a-b", "-", "_") = "a_b", logFile)

ExitApp(failures ? 1 : 0)

; 断言：失败时写入日志并累加计数
Check(name, condition, logFile) {
    global failures
    if !condition {
        failures += 1
        FileAppend "FAIL " name "`n", logFile
    }
}
```

```powershell
& "C:\AutoHotkey\v2\AutoHotkey64.exe" "W:\project\tests\run-tests.ahk"
if ($LASTEXITCODE -ne 0) { throw "测试失败，退出码：$LASTEXITCODE" }
```

### 4.6 打包与发布（需要时）

在入口脚本顶部加入 Ahk2Exe 元数据：

```autohotkey
; @Ahk2Exe-SetName MyProject
; @Ahk2Exe-SetVersion 0.1.0
; @Ahk2Exe-SetDescription MyProject
; @Ahk2Exe-SetMainIcon assets\app.ico
```

编译命令：

```powershell
& "C:\AutoHotkey\Compiler\Ahk2Exe.exe" `
  /in "W:\project\main_v0.1.0.ahk" `
  /out "W:\project\dist\MyProject_v0.1.0.exe" `
  /base "C:\AutoHotkey\v2\AutoHotkey64.exe" `
  /icon "W:\project\assets\app.ico"
```

发布前检查：

- [ ] 版本号、ChangeLog、元数据和产物文件名一致。
- [ ] 测试退出码为 0。
- [ ] 没有调试用 `MsgBox`、`OutputDebug`、开发热键、绝对开发路径或密钥。
- [ ] 在未安装 AutoHotkey 的干净环境运行编译产物。
- [ ] 编译后的 AHK 脚本可被提取，严禁嵌入密钥或敏感信息。
- [ ] 如对外分发，附 SHA256 校验值，并说明未签名程序的 SmartScreen 提示。

## 附录 A：《XX项目整体情况概览.md》模板

````markdown
# <项目名> 整体情况概览

> 最后更新：YYYY-MM-DD HH:mm
> 当前版本：vX.Y.Z
> 当前文件：`W:\project\main_vX.Y.Z.ahk`

## 1. 项目目标与范围

- 目标：
- 核心功能：
- 明确不做：

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 示例模块 | 已完成 / 进行中 / 待启动 | |

## 3. 已完成内容

- 

## 4. 未完成/待修改内容

- [ ] 

## 5. 下一步具体修改点

1. 文件：`W:\project\main_vX.Y.Z.ahk`
   - 函数/模块：
   - 当前行为：
   - 目标行为：
   - 输入/输出：
   - 验收方式：

## 6. 关键文件与目录说明

| 路径 | 作用 | 备注 |
| --- | --- | --- |
| `W:\project\main_vX.Y.Z.ahk` | 当前入口 | |
| `W:\project\lib\` | 共享库 | |

## 7. 运行与验证命令

```powershell
& "C:\AutoHotkey\v2\AutoHotkey64.exe" "W:\project\main_vX.Y.Z.ahk"
```

- 预期结果：
- 实际结果：
- 退出码：

## 8. 已知问题、风险与注意事项

- 

## 9. 最近变更记录

### vX.Y.Z（YYYY-MM-DD）

- 新增：
- 修改：
- 修复：
- 风险：

## 10. 给下一个 AI 账号的接续指令

- 先读取本文件，再复述当前进展和停留节点。
- 从 `文件:行号` 的 `函数名` 继续。
- 下一步只做：
- 不要做：
- 验证命令：
````

## 附录 B：执行前自检清单

- [ ] 已用中文给出方案，且未创建或修改任何文件。
- [ ] 已收到明确的“可以编写代码”授权。
- [ ] 已在动代码前更新《XX项目整体情况概览.md》，10 项齐全，路径和命令为绝对路径。
- [ ] 已创建递增版本号的新文件，旧版本仍然保留。
- [ ] 新文件头部有中文 ChangeLog，所有函数和核心分支有中文注释。
- [ ] 已用绝对路径运行 AHK，并记录实际结果和退出码。
- [ ] 若是代码修改，最后一行只有 `已完成`；若不是代码修改，没有该行。