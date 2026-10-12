# 作业辅导考勤次数统计（桌面 App）

把一份或多份考勤表加起来，统计每个学生在所有日期里出现的次数，汇总成一张
`学生姓名 / 行政班 / 需要补考勤次数` 的表格。

## 直接使用

**macOS**：双击 `AttendanceTally.app`。
首次打开若提示「无法打开，因为来自身份不明的开发者」，用右键 → 打开，
或在终端执行一次：

```bash
xattr -dr com.apple.quarantine /path/to/AttendanceTally.app
```

**Windows**：把 `AttendanceTally.exe` 拷过去双击即可（单文件，免安装）。
首次运行 SmartScreen 可能拦截，点「更多信息」→「仍要运行」。

## 用法

1. 点「添加文件...」选择考勤表，也可以把文件直接拖进列表。可以一次加多个。
2. 需要的话填「输出文件」路径；留空就放在第一个文件旁边，名为 `考勤次数统计.xlsx`。
3. 一个人有错别字写法时，在「合并同名」里每行写一条，例如 `尹湉黙=尹湉默`。
4. 点「开始统计」。下面的结果区会显示每个文件的人数/次数、合计、以及校验结果。

多个文件会**一起统计到同一张表**，按（姓名 + 班级）合并计数。

## 识别规则

- 表头里要有 `学生所属班级` 列，以及表头含 `学生姓名` 的名单列（每列一天）。
  有 `学生所属年级` 就一起用来自动拼出「六年级10班」这种行政班名。
- 一个 sheet 里可以上下堆叠多个班级，每段前面有各自表头，会自动识别。
- 班级列不是数字的行会被跳过，所以重复的表头行和结尾的老师行不会计入。
- 若同一学生同一天在多个输入文件里都出现，会提示重复计数（多半是两份表内容重叠）。

## 从源码运行 / 重新打包

需要 Python 3.9+ 且带 tkinter。

```bash
pip install -r requirements.txt
python app.py                    # 直接运行
python app.py --selftest 表.xlsx -o 结果.xlsx --merge 尹湉黙=尹湉默   # 无界面自检
```

打包：

- macOS：`bash build_macos.sh` → `dist/AttendanceTally.app`
- Windows：在 **Windows 机器上**运行 `build_windows.bat` → `dist\AttendanceTally.exe`
  （PyInstaller 不能跨系统打包，mac 上无法直接生成 exe）
- 不想配环境：推送到 GitHub 后跑 `.github/workflows/build.yml`，两个系统的产物会作为
  artifact 上传。

## 说明

- 统计逻辑与 `attendance-tally` 技能里的 `tally_attendance.py` 是同一套，因此结果一致。
- 默认统计完会用两种独立解析（openpyxl 和直接读工作表 XML）各算一遍并回读输出文件比对，
  任何不一致都会报错。不想要就取消勾选「统计后自动校验」。
- 打包出来的 app 自带 Python 和依赖，目标机器不需要装 Python。
- 想改中文名，直接把 `AttendanceTally.app` / `AttendanceTally.exe` 改名就行。
