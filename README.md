# FkRain
> 仅供学习与研究使用，请遵守法律法规与平台条款。请勿用于任何违规用途。


## 环境要求
- Windows PowerShell 5.1 或 PowerShell 7+（跨平台可用：Windows/macOS/Linux 需安装 PowerShell 7）

若遇到“运行脚本已被系统禁用”的提示，可在当前终端临时放开限制：
```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```


## 获取 Cookie
1. 登录雨课堂网页端：https://www.yuketang.cn/v2/web/index
2. 打开浏览器开发者工具（F12），在“网络/Network”面板随便点开一个请求。
3. 复制请求头里的 Cookie 字段，至少需要包含：
   - csrftoken=xxxxxxxx...;
   - sessionid=xxxxxxxx...;

脚本会从 Cookie 中自动尝试提取 csrftoken、uv_id/university_id 等信息；一般只填 csrftoken 和 sessionid 即可。

也可以通过环境变量提供：
```
$env:YKT_COOKIE = "csrftoken=...; sessionid=...;"
```


## 快速开始（交互式菜单，推荐）
在脚本所在目录打开 PowerShell：
```
# Windows PowerShell
.\new.ps1

# 或 PowerShell 7 跨平台
pwsh ./new.ps1
```
首次运行会显示免责声明与 Cookie 提示，按说明粘贴 Cookie 后：
- 选择“课程/班级”
- 在“回放列表”选择：
  - 自动刷本班全部回放（默认刷全部，不仅限未观看）
  - 手动选择单个回放
- 控制台会显示批量进度与心跳发送进度，完成后给出统计。


## 直连模式（单节回放）
当你已知 classroom_id 与 lesson_id 时，可不开菜单直接执行：
```
# 最小示例：自动补全其它必要字段（推荐）
.\new.ps1 -Cookie "csrftoken=...; sessionid=...;" -ClassroomId 123456 -LessonId 654321 -NoMenu
```
脚本会从接口自动补齐 user_id、course_id、video_id 与 duration。

若你已掌握全部参数，也可关闭自动补全：
```
.\new.ps1 -Cookie "csrftoken=...; sessionid=...;" \
  -UserId 10001 -CourseId 20002 -ClassroomId 123456 -LessonId 654321 \
  -VideoId 30003 -Duration 3600 -Interval 5 -BatchSize 120 \
  -Timeout 20 -Retries 2 -SleepMsBetweenBatch 80 \
  -NoMenu -NoAutoFill
```

调试查看不实际发送（只生成心跳并打印前几条）：
```
.\new.ps1 -Cookie "csrftoken=...; sessionid=...;" -ClassroomId 123456 -LessonId 654321 -NoMenu -DryRun
```


## 参数说明（摘自脚本）
- `-Cookie <string>`：登录 Cookie；若未提供，优先读取环境变量 `YKT_COOKIE`，否则运行时提示粘贴。
- `-UserId <int?>`：用户 ID。留空时在自动补全开启下从接口获取。
- `-CourseId <string>`：课程 ID。留空时在自动补全开启下从接口获取。
- `-ClassroomId <string>`：班级 ID。
- `-LessonId <string>`：回放（课节）ID。
- `-VideoId <string>`：视频 ID。留空时在自动补全开启下从接口获取或从回放信息中择优选择。
- `-Duration <double?>`：视频总时长（单位：秒）。留空时在自动补全开启下从接口获取。
- `-Interval <int>`（默认 5）：心跳间隔（秒）。
- `-BatchSize <int>`（默认 120）：每次 POST 的心跳条数；<=0 表示一次性发送全部。
- `-Timeout <int>`（默认 20）：HTTP 请求超时（秒）。
- `-Retries <int>`（默认 2）：失败重试次数（按批次）。
- `-SleepMsBetweenBatch <int>`（默认 80）：两批心跳之间的休眠（毫秒），用于限速与避嫌。
- `-DryRun`：仅生成并展示心跳，不实际发送。
- `-AutoFixMs` / `-NoAutoFixMs`：是否自动将“看似毫秒”的 `Duration` 纠正为秒（默认开启）。
- `-AutoFill` / `-NoAutoFill`：是否从接口自动补全缺失字段（默认开启）。
- `-Menu` / `-NoMenu`：是否启用交互式菜单（默认开启）。

直连模式下，至少需要：`-ClassroomId` 与 `-LessonId`。其余可由 `-AutoFill` 自动补齐。


## 使用提示
- 批量模式会在每节完成后默认暂停 10 秒；心跳分批发送时也会按 `-SleepMsBetweenBatch` 间隔休眠。
- 选择视频源时，脚本会自动挑选“时长较长且排序更优”的一条作为心跳目标。
- 控制台进度条依赖终端宽度，终端太窄可能显示不全。建议使用 Windows Terminal 或等宽字体终端。


## 常见问题（FAQ）
- 401/403 或接口报错：Cookie 过期或无效。请重新登录并复制最新 Cookie。
- 课程/回放列表为空：账号无该课程或班级；或当前大学/学期无可见数据。
- 运行脚本被禁止：使用 `Set-ExecutionPolicy -Scope Process Bypass` 临时解除限制。
- 进度条“错位/花屏”：调整终端宽度，或在较新的终端程序中运行。


## 免责声明
- 本工具仅用于学习与研究。
- 使用者应自行承担一切风险与后果，包括但不限于账号风险、数据风险等。
- 脚本作者与仓库维护者不对任何不当使用负责。


## 致谢
- PowerShell 社区与相关开源工具。
