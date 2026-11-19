---
name: solidworks
description: 用本机 SOLIDWORKS 2025 直接建模：新建零件/装配/工程图、草图、拉伸/切除/旋转、圆角/倒角、导出 STEP/STL/IGES/PDF、读质量属性与包围盒。当用户要求"在 SolidWorks 里做个零件/改模型/出工程图/导出模型"时使用。
---

# SolidWorks 接入（本机 SW2025 COM 自动化）

桥接脚本：`swbridge.py`，入口：`sw.cmd` / `sw.ps1`（自动挑一个装了 pywin32 的 Python）。
所有命令输出一行 JSON，长度单位统一为 **毫米（mm）**，脚本内部换算成米。

调用方式（在仓库根目录）：

```powershell
.\sw.cmd <命令> [--参数 值 ...]
```

## 前提

- 机器上已装 SOLIDWORKS 2025（`D:\SolidWorks 2025\SOLIDWORKS\SLDWORKS.exe`），COM 注册正常。
- Python 需要 `pywin32`：`py -3 -m pip install -r requirements.txt`。
- 第一次调用会启动 SolidWorks（可能要 1–3 分钟，若返回 `CO_E_SERVER_EXEC_FAILURE` 就等它启动完再重试，或用 `Start-Process 'D:\SolidWorks 2025\SOLIDWORKS\SLDWORKS.exe'` 先拉起）。
- 零件模板用 `C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\gb_part.prtdot`（自动查找）。

## 常用命令

| 目的 | 命令 |
|---|---|
| 看当前状态 | `status` |
| 新建零件/装配/工程图 | `new --type part` / `new --type assembly` / `new --type drawing` |
| 打开已有文档 | `open "D:\path\part.SLDPRT"` |
| 画草图 | `sketch --plane front --shapes "rect 0 0 60 40"` |
| 拉伸 | `extrude --depth 10`（`--through-all`、`--mid`、`--draft 3`、`--flip`） |
| 切除 | `cut --dir`（默认贯穿；`--blind --depth 5`；方向不对加 `--dir` / `--flip`） |
| 旋转 | `revolve --angle 360`（草图里要有中心线/轴线） |
| 圆角 | `fillet --radius 3 --edges vertical` |
| 倒角 | `chamfer --distance 1 --edges vertical` |
| 保存 | `save --path "out\plate.sldprt"` |
| 导出 | `export --out "out\plate.step"`（step/stl/igs/pdf 都行） |
| 质量属性/包围盒/特征树 | `info` |
| 截图 | `screenshot --out "out\view.bmp"` |
| 跑 VBA 宏 | `macro "C:\path\m.swp"` |
| 任意 COM 代码兜底 | `exec --code "..."`（预置 `sw`、`model`、`bridge`、`MM`、`val`） |

### 草图 DSL（`--shapes`，分号或换行分隔，全为 mm）

```
rect cx cy w h            中心 + 宽高
circle cx cy r            圆心 + 半径
line x1 y1 x2 y2
polygon cx cy r n [rot]   正 n 边形
slot cx cy length width [angle]   槽（三条线近似）
```

`--shapes-file file.txt` 可从文件读；`--plane` 支持 `front|top|right`，也支持按名字模糊匹配（界面是中文时就是 `前视基准面` 等，脚本按"第 1/2/3 个基准面"定位，不受语言影响）。

### 选边（圆角/倒角）

`--edges` 取值：

- `vertical` — 平行 Z 轴的边（板类零件立边，最常用）
- `horizontal` — 垂直于 Z 轴的直线边
- `circular` — 圆/相切边
- `all` — 全部边
- `near:30,20,5;0,0,0` — 按坐标点就近选边（mm）

## 标准工作流

1. `new --type part`
2. `sketch --plane front --shapes "rect 0 0 80 50"` → `extrude --depth 12`
3. 打孔：`sketch --plane front --shapes "circle 20 15 4; circle -20 15 4"` → `cut --dir`
4. 圆角：`fillet --radius 2 --edges vertical`
5. `info` 核对体积/外形 → `save --path ...` → `export --out ...`

每步都会返回 `"ok": true/false`，失败时带 `error`。**每次改完都用 `info` 核对 `bbox_mm.size` 和 `volume_mm3` 是否符合预期**，不要凭"命令没报错"就下结论。

## 三个真实坑（已在 SW2025 上验证过）

1. **切除方向**：草图在前视基准面上、实体往 +Z 拉伸时，`cut` 默认朝 -Z 会切空，返回 `cut failed`。加 `--dir`（或 `--flip`）朝实体那一侧切。
2. **圆角参数**：内部 `FeatureFillet3` 的 Options 必须是 `2`，边上要先选中；`--edges vertical` 已封装好。
3. **零参数方法是属性**：SolidWorks 的 IDispatch 把 `GetTitle`/`EditRebuild3`/`GetStartVertex` 这类零参方法暴露成属性，直接 `obj.Method()` 会报"找不到成员"。脚本里的 `val(obj, "Name")` 已统一处理，写 `exec` 兜底代码时也要用它。

## 自定义/扩展

- 需要桥接脚本没有的功能时，用 `exec` 直接写 COM 调用，例如：
  ```powershell
  .\sw.cmd exec --code "m=model; print(m.GetTitle); print([f.Name for f in bridge.features()])"
  ```
  复杂几何建议先用 `exec` 验证参数，再把常用套路补进 `swbridge.py`。
- 已有的 SolidWorks 宏可直接 `macro` 运行。

## 已知限制

- 目前没有装配体配合、工程图视图、方程式/配置驱动这些高层命令；要这些用 `exec` 或告诉开发者补命令。
- 曲面上选面/选边只支持边，不支持按面选。
- UI 是中文界面，特征名会是 `凸台-拉伸1`、`切除-拉伸1`、`圆角1` 这类；按名字引用时用 `--sketch 草图2` 这种前缀即可。