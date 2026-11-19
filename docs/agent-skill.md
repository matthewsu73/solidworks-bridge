---
name: solidworks
description: 用本机 SOLIDWORKS 2025 直接建模：新建零件/装配/工程图、草图、拉伸/切除/旋转、圆角/倒角、曲面拉伸/旋转/放样/扫描/延伸/偏移/缝合/加厚、导出 STEP/STL/IGES/PDF、读质量属性与包围盒。当用户要求"在 SolidWorks 里做个零件/改模型/出工程图/导出模型"时使用。
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
| 偏移基准面 | `plane-offset --base front --distance 40`（返回名字，如 `基准面1`，可直接给 `--plane`） |
| 平面区域曲面 | `surface-plane` |
| 拉伸曲面 | `surface-extrude --depth 25`（开口薄壳，不是实体） |
| 旋转曲面 | `surface-revolve --angle 360`（草图里要有 `centerline`） |
| 放样曲面 | `surface-loft --sketches 草图1,草图2` |
| 扫描曲面 | `surface-sweep --outline 草图1 --path 草图2` |
| 加厚成实体 | `thicken --thickness 3 --faces all|near:x,y,z` |
| 偏移曲面 | `offset-surface --distance 5 --faces all|near:x,y,z` |
| 缝合曲面 | `sew --tolerance 0.01` |
| 延伸曲面 | `extend-surface --distance 10 --edges all` |
| 实体布尔 | `combine --operation subtract`（`add|subtract|intersect`，`--target largest|smallest|#i|near:x,y,z`，`--tools all|...`） |

### 草图 DSL（`--shapes`，分号或换行分隔，全为 mm）

```
rect cx cy w h            中心 + 宽高
circle cx cy r            圆心 + 半径
line x1 y1 x2 y2
polygon cx cy r n [rot]   正 n 边形
slot cx cy length width [angle]   槽（三条线近似）
centerline x1 y1 x2 y2            构造线/旋转轴（旋转、旋转曲面必需）
```

`--shapes-file file.txt` 可从文件读；`--plane` 支持 `front|top|right`，也支持按名字模糊匹配（界面是中文时就是 `前视基准面` 等，脚本按"第 1/2/3 个基准面"定位，不受语言影响）。

### 选边（圆角/倒角）

`--edges` 取值：

- `vertical` — 平行 Z 轴的边（板类零件立边，最常用）
- `horizontal` — 垂直于 Z 轴的直线边
- `circular` — 圆/相切边
- `all` — 全部边
- `near:30,20,5;0,0,0` — 按坐标点就近选边（mm）

## 曲面

实体（solid）和曲面体（sheet）在 SOLIDWORKS 里是两个集合，**互不包含**：
`GetBodies2(0)` 只有实体，`GetBodies2(1)` 才有曲面。`info` 会给出
`solid_bodies`、`surface_bodies`，有曲面时还有 `surface_area_mm2`（逐面
`IFace2.GetArea` 累加，mm²）。开放曲面没有体积，`volume_mm3` 只算实体。

### 拉伸曲面 → 加厚成实体

```powershell
.\sw.cmd new --type part
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd surface-extrude --depth 25      # surface_area_mm2 = 5000
.\sw.cmd thicken --thickness 3 --faces all   # volume_mm3 = 15900
```

### 放样曲面（两个截面 + 偏移基准面）

```powershell
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd plane-offset --base front --distance 40    # -> "plane": "基准面1"
.\sw.cmd sketch --plane 基准面1 --shapes "circle 0 0 15"
.\sw.cmd surface-loft --sketches 草图1,草图2       # bbox 60x40x40
```

### 扫描曲面

```powershell
.\sw.cmd sketch --plane front --shapes "circle 0 0 5"
.\sw.cmd sketch --plane right --shapes "line 0 0 40 0"
.\sw.cmd surface-sweep --outline 草图1 --path 草图2  # bbox 10x10x40
```

`surface-sweep` 内部固定：轮廓草图 `Select2(True, 1)`、路径草图 `Select2(True, 4)`。
顺序反了可能“成功”但几何错误；路径还必须穿过轮廓所在平面。

### 平面曲面 → 延伸 / 偏移 / 缝合

```powershell
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd surface-plane
.\sw.cmd extend-surface --distance 10 --edges all   # 2400 -> 4800 mm2
.\sw.cmd offset-surface --distance 5 --faces near:0,0,0
.\sw.cmd sew                                        # 两个曲面体 → 一个
```

**选面不用点界面**：`--faces all` 全选，`--faces near:x,y,z`（mm）取三角面片质心
最接近该点的单个面。加厚用 mark 1、偏移曲面用 mark 0、缝合体用 `Body2.Select(True, 1)`，
这些标记已在 SW2025 上验证。

### 阳模 → 阴模（布尔求差）

阴模 = 料块 − 型芯实体。料块先拉伸成独立实体，型芯用曲面加厚得到，然后：

```powershell
.\sw.cmd sketch --plane front --shapes "rect 0 0 120 120"
.\sw.cmd extrude --depth 25                 # 料块 360000 mm3
# ...建阳模曲面，再 thicken 成实体...
.\sw.cmd combine --operation subtract --target largest
```

```jsonc
{ "ok": true, "feature": "组合1", "operation": "subtract",
  "target": "凸台-拉伸1", "target_volume_mm3": 360000.0,
  "tools": ["凸台-拉伸2"], "volume_before_mm3": [12566.37, 360000.0],
  "volume_mm3": 347433.63, "solid_bodies": 1 }
```

`--target` 默认 `largest`（体积最大的实体，通常就是料块）；`--tools all` 减掉其余全部实体。
布尔只能作用于**实体**，所以型芯必须先 `thicken`。

## 标准工作流

1. `new --type part`
2. `sketch --plane front --shapes "rect 0 0 80 50"` → `extrude --depth 12`
3. 打孔：`sketch --plane front --shapes "circle 20 15 4; circle -20 15 4"` → `cut --dir`
4. 圆角：`fillet --radius 2 --edges vertical`
5. `info` 核对体积/外形 → `save --path ...` → `export --out ...`

每步都会返回 `"ok": true/false`，失败时带 `error`。**每次改完都用 `info` 核对 `bbox_mm.size` 和 `volume_mm3` 是否符合预期**，不要凭"命令没报错"就下结论。

## 几个真实坑（已在 SW2025 上验证过）

1. **切除方向**：草图在前视基准面上、实体往 +Z 拉伸时，`cut` 默认朝 -Z 会切空，返回 `cut failed`。加 `--dir`（或 `--flip`）朝实体那一侧切。
2. **圆角参数**：内部 `FeatureFillet3` 的 Options 必须是 `2`，边上要先选中；`--edges vertical` 已封装好。
3. **零参数方法是属性**：SolidWorks 的 IDispatch 把 `GetTitle`/`EditRebuild3`/`GetStartVertex` 这类零参方法暴露成属性，直接 `obj.Method()` 会报"找不到成员"。脚本里的 `val(obj, "Name")` 已统一处理，写 `exec` 兜底代码时也要用它。
4. **创建特征的 API 常返回 `None`**：成功与否要看特征树有没有新增项，
   不能凭返回值判断。
5. **选面标记因命令而异**：加厚 `1`、偏移曲面 `0`、缝合体 `1`；标错时
   往往不报错也不生成特征。
6. **布尔要传 VARIANT 数组**：`InsertCombineFeature(15902, commonBody, bodies)` 第三参必须是
   `win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [...])`，第二参也不能是 `None`。

## 自定义/扩展

- 需要桥接脚本没有的功能时，用 `exec` 直接写 COM 调用，例如：
  ```powershell
  .\sw.cmd exec --code "m=model; print(m.GetTitle); print([f.Name for f in bridge.features()])"
  ```
  复杂几何建议先用 `exec` 验证参数，再把常用套路补进 `swbridge.py`。
- 已有的 SolidWorks 宏可直接 `macro` 运行。

## 已知限制

- 目前没有装配体配合、工程图视图、方程式/配置驱动这些高层命令；要这些用 `exec` 或告诉开发者补命令。
- 曲面裁剪（剪裁曲面 / `PreTrimSurface`）暂未封装，需要时在界面里做或用 `exec` 试；
  也没有边界曲面（Boundary Surface）。
- 选面只有 `all` 和 `near:x,y,z`，没有按面索引/名称选择；选边可按几何特征。
- UI 是中文界面，特征名会是 `凸台-拉伸1`、`切除-拉伸1`、`圆角1` 这类；按名字引用时用 `--sketch 草图2` 这种前缀即可。