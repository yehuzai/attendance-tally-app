# AGENTS.md

作业辅导考勤次数统计：一个桌面 App，把一份或多份考勤表汇总成一张
`学生姓名 / 行政班 / 需要补考勤次数` 表格。使用者通常不懂代码，改动要保守。

## 文件分工

- `app.py`：Tkinter 界面 + 打包入口；带 `--selftest` 可无界面跑完整流程
- `attendance_tally.py`：全部统计逻辑（读取、计数、排序、写出、校验）
- `build_macos.sh`、`build_windows.bat`：PyInstaller 打包
- `.github/workflows/build.yml`：push 到 main 自动构建 mac / windows 产物

## 硬性约束

- `attendance_tally.py` 与技能 `attendance-tally` 里的
  `~/.agents/skills/attendance-tally/scripts/tally_attendance.py` 是同一份逻辑。
  **改任何一边都要同步另一边**，否则命令行的结果和 app 的结果会不一致。
- 绝不提交学生数据：`.gitignore` 已挡 `*.xlsx` `*.xlsm` `*.xls`。
- 不提交构建产物：`*.app/`、`build/`、`dist/`、`.venv/` 都已忽略。
- PyInstaller 不能跨系统打包：mac 上只能出 `.app`，`.exe` 必须靠 Windows 机器或 CI。
- macOS 产物是 arm64（Apple 芯片）；Intel Mac 需要自行打包。

## 统计规则（不要随手改）

- 表头含 `学生所属班级` 的列是班级列；表头含 `学生姓名` 的列是每天的名单列。
- 班级列不是数字的行一律跳过，这样重复的表头行和结尾的老师行会自动排除。
- 每个非空姓名单元格算一次，含同一人在同一天出现两次。
- 行政班名由 `<年级><班级>班` 拼出（`学生所属年级` 可选）。
- 多份输入按（姓名 + 班级）合并计数，统一输出到一张表。

## 改完必须验证

跑自检，并核对已知基线没有被改坏：

```bash
python app.py --selftest "六年级10班、11班第一、二周作业辅导考勤信息.xlsx" -o /tmp/out.xlsx --merge 尹湉黙=尹湉默
```

基线：**57 人 / 350 次**，其中 10 班 20 人 121 次、11 班 37 人 229 次；`尹湉默` 为 5 次；
输出里不应出现老师行。数字对不上就说明改坏了。

改界面时至少确认窗口能起来，且 `winfo_reqwidth/reqheight` 不超过设定的 geometry，
否则内容会被裁掉。

## 发布

代码 push 到 main 后 CI 会自动出两个 artifact，但 artifact 需要登录且 90 天过期。
对外发用 Release（mac 的 zip 要用 `ditto -c -k --keepParent AttendanceTally.app out.zip` 打，
这样才保得住 app 的可执行权限）：

```bash
gh release create v1.0.1 --repo yehuzai/attendance-tally-app --notes-file NOTES.md AttendanceTally-macos.zip AttendanceTally.exe
```
