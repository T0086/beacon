# schedule-orb（日程球）

Windows 桌面悬浮小组件：课程表 + 备忘 + 系统指标。PowerShell + WPF 单文件程序，不编译、不装运行时。

![preview](docs/preview.png)

## 有什么

**收起态**（470×76 圆角长条）：

```
11:56
15 · 周日    CPU 9%   67°C   GPU 26%   Research Methods    🔓 📌
             RAM 14.9/31.1G   6°C 阴    明天 09:00
```

- 时间 / 日期 / 周几（三行，按醒目度排）
- CPU 占用、CPU 温度、GPU 占用、RAM、天气
- **下一件事**：最近的课程或 deadline，deadline 前有绿/黄/红小球表示紧迫度
- 锁定按钮（防误触）、层级切换（置顶 / 沉底）

**展开面板**：

- **日程**：接下来 5 个上课日，按天分组
- **月历**：点击日期画红圈，下方显示那天日程
- **备忘**：输入框 + 日期选择，带日期的备忘自动挂到对应那天的课程下
- **设置**：粘贴 `.ics` 链接导入外部日历，每天 22:00 自动刷新

## 依赖

**零安装依赖。** 所有东西都在包里：

| 组件 | 用途 | 说明 |
|---|---|---|
| PowerShell 5.1 | 运行环境 | Windows 10/11 自带 |
| [LibreHardwareMonitorLib](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor) **0.9.3** | CPU/GPU 温度 | 已内置于 `app/lib/`，MPL 2.0 |
| `nvidia-smi` | GPU 占用 | NVIDIA 驱动自带，没有就显示 `--` |
| [Open-Meteo](https://open-meteo.com) | 天气 | 免费接口，无需 key |

> **为什么用 0.9.3 而不是最新版？** 0.9.6 开始依赖 PawnIO 内核驱动，装起来要跟杀软、驱动签名、UAC 三重门纠缠。0.9.3 用内嵌的 WinRing0 驱动，运行时自动加载，**不需要安装任何驱动**。

## 安装

1. 解压整个压缩包
2. 双击 **`安装.bat`**
3. UAC 弹窗点「是」

安装脚本会：

1. 复制 `app/` 到 `%LOCALAPPDATA%\schedule-orb\`
2. 建计划任务 `Schedule Orb`（登录时以**最高权限**启动，不弹 UAC）
3. 移除可能存在的旧启动文件夹快捷方式
4. 启动

> **为什么要最高权限？** WinRing0 驱动要读 CPU 的 MSR 寄存器，普通权限下 `Computer.Open()` 能成功但读不到传感器。用计划任务提权可以在不弹 UAC 的前提下拿到权限。

**卸载**：删除计划任务 `Schedule Orb` + 目录 `%LOCALAPPDATA%\schedule-orb\`。

## 数据文件

| 文件 | 内容 |
|---|---|
| `notes.json` | 备忘（含 `due` 字段） |
| `imported.json` | 导入的 ics 日历源 |
| `state.json` | 窗口位置 |

课程数据默认读 `..\desktop-calendar\schedule.json`。**没有这个文件也能跑**——只是日程区显示「没有更多课程」，可以靠「设置」里导入 `.ics` 链接补上。

## 已知限制

- 中文输入法下回车不提交备忘（IME 吃掉了），所以主入口是 `+` 按钮
- 需要管理员权限（计划任务），否则温度显示 `--°C`
- 多屏且缩放不同时，鼠标命中可能有偏移（已加 `SetProcessDPIAware` 缓解）
- 同一时间只允许一个实例（全局互斥量）

## 许可

本项目代码：MIT。

第三方组件：

- LibreHardwareMonitorLib：**MPL 2.0**（见 `LICENSES/`）
- WinRing0 驱动：随 LHM 分发

分发时请保留 `LICENSES/` 目录。

---

`docs/开发笔记.md` 记了一堆 PowerShell 5.1 + WPF 的坑，改之前值得先看。
