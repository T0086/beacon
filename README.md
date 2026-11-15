# schedule-orb（日程球）

Windows 桌面悬浮小组件：课程表 + 备忘 + 系统指标，用 PowerShell + WPF 写的单文件程序，不需要编译、不需要安装。

## 有什么

**收起态**（470×76 圆角长条）：

```
11:30
15 · 周日    CPU 13%  60°C  GPU 34%   Research Methods    🔓 📌
             RAM 15/31.1G   6°C 阴    明天 09:00
```

- 时间 / 日期 / 周几
- CPU 占用、CPU 温度、GPU 占用、RAM、天气
- **下一件事**：最近的课程或 deadline，deadline 前有绿/黄/红小球表示紧迫度
- 锁定按钮（防误触）、层级切换（置顶 / 沉底）

**展开面板**：

- **日程**：接下来 5 个上课日，按天分组，带时间、教室、课程配色
- **月历**：7×6 网格，点击日期画红圈，下方显示那天日程
- **备忘**：输入框 + 日期选择，带日期的备忘会自动挂到对应那天的课程下面
- **设置**（独立窗口）：粘贴 `.ics` 链接导入外部日历，每天 22:00 自动刷新

## 依赖

| 组件 | 用途 | 是否必需 |
|---|---|---|
| Windows 10/11 + PowerShell 5.1 | 运行环境 | 必需 |
| [LibreHardwareMonitorLib](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor) | 读 CPU 温度 | 可选 |
| [PawnIO](https://pawnio.eu) | LHM 的内核驱动 | 配合 LHM |
| AIDA64 | 另一条 CPU 温度路径（注册表导出） | 可选，二选一 |
| `nvidia-smi` | GPU 占用（NVIDIA 卡自带） | 可选 |
| [Open-Meteo](https://open-meteo.com) | 天气（免费，无需 key） | 可选 |

CPU 温度两条路径都兼容：有 AIDA64 读 `HKCU\Software\FinalWire\AIDA64\SensorValues`，有 LHM 就加载 DLL。都没有就显示 `--°C`。

## 安装

1. 双击 `启动 日程球.vbs`（不弹黑框）
2. 开机自启：把该 vbs 的快捷方式放进
   `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`

## 数据文件

| 文件 | 内容 |
|---|---|
| `notes.json` | 备忘（含 `due` 字段） |
| `imported.json` | 导入的 ics 日历源 |
| `state.json` | 窗口位置 |

课程数据读 `..\desktop-calendar\schedule.json`，由 `build.py` 从 `calendar.ics` 生成。

## 已知限制

- 中文输入法下回车不提交备忘（IME 吃掉了），所以主入口是 `+` 按钮
- 面板从长条左侧展开，靠屏幕左缘会被推回来
- 多屏且缩放不同时，鼠标命中可能有偏移（已加 `SetProcessDPIAware` 缓解）

## 许可

本项目代码：MIT。

第三方组件许可见 `LICENSES/`。分发时请保留：

- LibreHardwareMonitorLib：MPL 2.0
- PawnIO：GPL v2（建议以独立安装器形式随包分发，属 aggregation）

---

代码是长期迭代出来的，`docs/开发笔记.md` 里记了一堆 PowerShell 5.1 + WPF 的坑，改之前值得先看。
