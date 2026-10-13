把一份或多份考勤表汇总，统计每个学生在所有日期里出现的次数，输出一张
`学生姓名 / 行政班 / 需要补考勤次数` 的表格。自带 Python 和依赖，目标机器不用装任何东西。

## 本次更新（v1.0.1）

- 去掉「合并同名」这个手工预处理：不用再提前填 `尹湉黙=尹湉默` 之类的规则。
- 同一个班里出现两个只差一个字的名字时，程序不再自动合并，而是**各算各的次数**，
  并把这两行在结果表里用黄色标出，同时在结果区给出提示，由你自己判断是不是同一个人。
- 不同班级之间的相似名字不会被比较，也不会被标注。

## 下载哪个

| 系统 | 文件 | 用法 |
| --- | --- | --- |
| macOS（Apple 芯片 / M 系列） | `AttendanceTally-macos.zip` | 解压后得到 `AttendanceTally.app`，双击运行 |
| Windows 64 位 | `AttendanceTally.exe` | 直接双击，免安装 |

Intel 芯片的 Mac 目前没有现成包，需要照 README 在 Intel 机器上执行 `build_macos.sh` 自行打包。

## 首次打开被系统拦截怎么办

- macOS：这个 app 没有做苹果签名/公证。若提示「无法打开，因为来自身份不明的开发者」，
  右键点 app →「打开」；或先在终端跑一次
  `xattr -dr com.apple.quarantine AttendanceTally.app`。
- Windows：若 SmartScreen 拦截，点「更多信息」→「仍要运行」。

## 校验下载是否完整（SHA-256）

```
29795d697a7498e009b9941d7933e748694e952180b5783a8bf6af8d5071dfb6  AttendanceTally-macos.zip
56cf24ba26580e9bf89ed746f0bdcd365db0a18da68bd3d3e95e24260f978c09  AttendanceTally.exe
```

macOS 校验：`shasum -a 256 AttendanceTally-macos.zip`
Windows 校验：`certutil -hashfile AttendanceTally.exe SHA256`
