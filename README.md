# solidworks-bridge

**Drive SOLIDWORKS 2025 from the command line. One command, one JSON line, millimetres.**

A thin, late-bound COM bridge that lets scripts and coding agents build and inspect
SolidWorks parts without touching the UI. Every subcommand prints exactly one JSON
object, so the caller can verify each step instead of guessing from the screen.

> 非官方项目。与 Dassault Systèmes 无关，不使用任何官方 SDK，只是调用本机已安装的
> SOLIDWORKS 所提供的 COM 接口。请在你拥有合法授权的 SOLIDWORKS 上使用。

| | |
|---|---|
| 平台 | Windows |
| 目标 | SOLIDWORKS 2025（在 33.5.0 上验证；2021+ 大概率可用，未逐一测试） |
| 依赖 | Python 3.8+、`pywin32` |
| 单位 | 命令行一律毫米，内部换算成米 |
| 输出 | 每命令一行 JSON |
| 许可 | MIT |

![示例](docs/images/flange_demo.png)

## 为什么会有这个东西

SOLIDWORKS 没有官方 MCP，也没有跨平台的自动化接口。用界面自动化（截图 + 点菜单）
慢且脆；直接写 COM 脚本每次都要重新处理一堆坑（参数个数、按引用传值、零参方法被
暴露成属性……）。这个项目把那些坑填平，留一个稳定的命令行边界：

- 一步一条命令，返回结构化结果，失败带 `error` 与 traceback
- 每步都能用 `info` 核对体积和包围盒，不靠"看起来对"
- 全本地运行，不联网、不上传模型

## 安装

```powershell
git clone <this-repo>
cd solidworks-bridge
py -3 -m pip install -r requirements.txt
```

脚本会按顺序找 Python：`$env:SW_PYTHON` → 仓库内 `.venv` → `py -3` → `python3` → `python`，
并用第一个能 `import win32com.client` 的。想固定解释器：

```powershell
$env:SW_PYTHON = "C:\Python312\python.exe"
```

前提：本机装好 SOLIDWORKS 且 COM 已注册（正常安装即可）。首次调用如果 SolidWorks
还没启动，可能返回 `CO_E_SERVER_EXEC_FAILURE`——先把 SOLIDWORKS 打开，或重试一次。

## 快速开始

```powershell
.\sw.cmd status
.\sw.cmd new --type part
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd extrude --depth 10
.\sw.cmd sketch --plane front --shapes "circle 0 0 8"
.\sw.cmd cut --dir
.\sw.cmd fillet --radius 3 --edges vertical
.\sw.cmd info
.\sw.cmd save   --path out\plate.sldprt
.\sw.cmd export --out  out\plate.step
```

这段就是 `docs/images/flange_demo.png` 那个零件：60×40×10 板 + Ø16 通孔 + 4×R3 立边圆角，
体积 21912 mm³，解析值 21912.09（差 0.03 来自圆弧离散）。

## 命令

| 命令 | 说明 |
|---|---|
| `status` | 版本号、打开文档数、当前文档概览 |
| `new --type part\|assembly\|drawing` | 新建文档，可 `--template`、工程图 `--paper-size` |
| `open <path>` | 打开已有文档 |
| `save [--path <path>]` | 另存 / 保存 |
| `close` | 关闭当前文档（未保存会弹窗，先 `save`） |
| `sketch --plane front\|top\|right --shapes "..."` | 画草图，按名字模糊匹配基准面（中文界面也行） |
| `extrude --depth 10` | 拉伸：`--through-all` `--mid` `--draft 3` `--flip` |
| `cut` | 切除，默认贯穿：`--blind --depth 5` `--dir` `--flip` |
| `revolve --angle 360` | 旋转（草图里要有中心线/轴线） |
| `fillet --radius 3 --edges vertical` | 圆角 |
| `chamfer --distance 1 --edges vertical` | 倒角 |
| `plane-offset --base front --distance 40` | 建偏移基准面（`--distance` 可为负），返回名字（如 `基准面1`）可直接给 `--plane` |
| `surface-plane` | 用草图生成平面区域曲面 |
| `surface-extrude --depth 25` | 拉伸曲面（开口薄壳，不是实体） |
| `surface-revolve --angle 360` | 旋转曲面（草图里要有中心线） |
| `surface-loft --sketches 草图1,草图2` | 放样曲面 |
| `surface-sweep --outline 草图1 --path 草图2` | 扫描曲面 |
| `thicken --thickness 3` | 曲面加厚成实体，`--faces all\|near:x,y,z`，`--no-merge` 让结果独立成体 |
| `offset-surface --distance 5` | 等距偏移曲面，`--faces all\|near:x,y,z` |
| `sew` | 缝合曲面体，`--tolerance 0.01` |
| `extend-surface --distance 10` | 延伸曲面边，`--edges all\|near:...` |
| `combine --operation subtract` | 实体布尔：`add\|subtract\|intersect`，`--target largest\|smallest\|#i\|near:x,y,z` |
| `export --out out\x.step` | 导出 step / stl / igs / pdf |
| `info` | 包围盒、体积、面积、质心、实体/曲面体数量、曲面面积、特征树 |
| `rebuild` | 重建并回读体积 |
| `screenshot --out out\v.bmp` | 存视图截图 |
| `macro <path> [--module --method]` | 跑已有 VBA 宏 |
| `exec --code "..."` | 任意 COM 兜底，预置 `sw`、`model`、`bridge`、`MM`、`val` |

### 草图 DSL

`--shapes`（分号或换行分隔，全部毫米），或 `--shapes-file file.txt`（`-` 读 stdin）：

```
rect    cx cy w h
circle  cx cy r
line    x1 y1 x2 y2
polygon cx cy r n [rot]
slot    cx cy length width [angle]
centerline x1 y1 x2 y2  构造线 / 旋转轴（旋转、旋转曲面用）
```

### 选边（圆角 / 倒角）

`vertical`（平行 Z 轴的边，板类零件最常用）、`horizontal`、`circular`（圆/相切边）、
`all`、`near:30,20,5;0,0,0`（按坐标就近，毫米）。

## 曲面建模

曲面体（sheet body）和实体（solid body）在 SOLIDWORKS 里是两个独立集合。`info` 会分别报
`solid_bodies` / `surface_bodies`；开放曲面没有有意义的体积，所以面积改用
`IFace2.GetArea` 逐面累加得到 `surface_area_mm2`，`volume_mm3` 只统计实体。

### 拉伸曲面 + 加厚（做成实体）

```powershell
.\sw.cmd new --type part
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd surface-extrude --depth 25
.\sw.cmd info            # surface_bodies 1, surface_area_mm2 5000
.\sw.cmd thicken --thickness 3 --faces all
.\sw.cmd info            # solid_bodies 1, volume_mm3 15900
```

### 放样曲面（需要偏移基准面）

```powershell
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd plane-offset --base front --distance 40     # 返回 "plane": "基准面1"
.\sw.cmd sketch --plane 基准面1 --shapes "circle 0 0 15"
.\sw.cmd surface-loft --sketches 草图1,草图2        # bbox 60x40x40
```

### 扫描曲面

```powershell
.\sw.cmd sketch --plane front --shapes "circle 0 0 5"
.\sw.cmd sketch --plane right --shapes "line 0 0 40 0"
.\sw.cmd surface-sweep --outline 草图1 --path 草图2  # bbox 10x10x40
```

轮廓和路径的选择标记是固定的（轮廓 mark 1、路径 mark 4），写反了会“成功”但几何是错的；
路径还必须与轮廓所在平面相交。

### 平面曲面 + 延伸 / 偏移 / 缝合

```powershell
.\sw.cmd sketch --plane front --shapes "rect 0 0 60 40"
.\sw.cmd surface-plane
.\sw.cmd extend-surface --distance 10 --edges all    # 2400 -> 4800 mm2
.\sw.cmd offset-surface --distance 5 --faces near:0,0,0
.\sw.cmd sew                                         # 两个曲面体合成一个
```

选面统一用 `--faces all` 或 `--faces near:x,y,z`（毫米，取三角面片质心最接近的那个面），
不依赖界面里点选。完整可运行脚本见 `examples\surface.ps1`。

### 阳模 → 阴模（布尔求差）

阴模 = 料块 − 型芯。料块先拉伸成独立实体；型芯（曲面加厚得到的实体）也必须是**另一个实体**，
再用 `combine subtract` 减掉。

```powershell
# 1) 120x120x25 料块
.\sw.cmd new --type part
.\sw.cmd sketch --plane front --shapes "rect 0 0 120 120"
.\sw.cmd extrude --depth 25                            # 360000 mm3

# 2) 型芯：曲面必须加厚成实体；--no-merge 让它保持独立（不加会直接并进料块）
.\sw.cmd thicken --thickness 3 --faces all --no-merge

# 3) 从料块里减掉型芯
.\sw.cmd combine --operation subtract --target largest
.\sw.cmd info      # volume_mm3 变小，solid_bodies 回到 1
```

已验证整条链路（`out\mould_cavity_demo.SLDPRT`）：料块 120×120×25，放样曲面
80×80(z=3) → Ø50(z=20)，加厚 3，减除后 **341300.26 mm³，回到 1 个实体**。

两个坑：

- 加厚壳的边别正好贴在料块表面上。完全共面时 SW 会切不干净——实测会把料块切成两块、
  或者留下碎料实体。壳体整块落在料块内部最稳；要开腔通到表面，用**实体**阳模做布尔更可靠。
- `combine subtract` 只认实体，曲面体必须先 `thicken`。

`--operation add` 合并、`intersect` 取交集；`--target` 默认取体积最大的实体（料块），
`--tools all` 表示减掉其余所有实体。要看是哪几个实体参与，先跑一次 `info`。

## 返回值怎么读

```jsonc
{
  "ok": true,
  "bbox_mm":       { "min": [...], "max": [...], "size": [...] },  // API 值，曲面上会偏大
  "bbox_mm_exact": { "min": [...], "max": [...], "size": [...] },  // 三角面片实测，可信
  "bodies": 1,
  "solid_bodies": 1,
  "surface_bodies": 0,          // 有曲面体时才会出现 surface_area_mm2
  "mass_properties": { "volume_mm3": 316526.79, "area_mm2": 42299.95, ... }
}
```

- `volume_mm3` / `area_mm2` 精确，可以直接拿来对理论值。
- `bbox_mm` 来自 `GetPartBox`，对含曲面的实体只是**近似**：一个真实的 120×25×120 零件
  可能报 124.5×27.9×124.5。要准的用 `bbox_mm_exact`（三角面片测量，误差 <0.01 mm，
  且因为是内接三角面，会**略微偏小**）。

## 踩过的坑

这些都是在本机 SW2025 上实测确认的，改代码前值得先看：

1. **类工厂不返回 `IDispatch`**（`E_NOINTERFACE`），C# 早期绑定互操作走不通；只能用
   pywin32 晚绑定 `win32com.client.Dispatch("SldWorks.Application")`。改过 interop 定义
   之后**必须**清掉过期的 `gen_py` 缓存。
2. **`OpenDoc6` 的 errors/warnings 是按引用输出参数**，直接传 `0` 会得到
   `com_error (-2147352571, '类型不匹配')`。要传
   `win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)`。
3. **参数个数比老文档多**：晚绑定下 `FeatureExtrusion3` 要 **23** 个参数、`FeatureCut4`
   要 **27** 个。少给会直接拒绝。
4. **`FeatureFillet3` 的 Options 必须是 `2`**，且边要先选中。
5. **零参方法被暴露成属性**：`GetTitle`、`EditRebuild3`、`GetStartVertex` 直接 `()` 调用
   会报"找不到成员"。脚本里统一用 `val(obj, "Name")`。
6. **切除方向**：草图在前视基准面、实体往 +Z 拉伸时，`cut` 默认朝 −Z 会切空；加 `--dir`。
7. **中性格式导出**用 `ModelDoc2.SaveAs3`；`Extension.SaveAs` 的 `ExportData` 传 `None`
   会类型不匹配。
8. **Windows 控制台默认不是 UTF-8**，中文特征名会变乱码并让 JSON 失效；脚本已强制
   `reconfigure(encoding="utf-8")`。
9. **`GetBodyBox` / `GetPartBox` 对曲面是近似值**，别拿它判断零件是否超料，用三角面片。
10. **实体和曲面体是两个集合**：`GetBodies2(0)` 只给实体，`GetBodies2(1)` 才给曲面，
    要一起统计就各查一次再合并。
11. **不少“创建特征”的 API 成功时返回 `None`**，不能拿返回值判成败；脚本改成比较调用前后
    特征树有没有新增项。
12. **布尔要传 VARIANT 数组**：`InsertCombineFeature(15902, commonBody, bodies)` 的第三个参数必须是
    `VARIANT(VT_ARRAY | VT_DISPATCH, [...])`，直接传 Python list 会报"类型不匹配"；
    而且第二个参数（保留实体）`None` 也不行，必须给 Body2。
13. **选面标记因命令而异**：加厚是 `Select2(True, 1)`，偏移曲面是 `Select2(True, 0)`，
    缝合体是 `Body2.Select(True, 1)`。标错时可能不报错但什么都不生成。

## 限制

- 装配体配合、工程图视图、方程式/配置驱动：没有高层命令，用 `exec` 写 COM，或自行补
  `swbridge.py`。
- 曲面：已支持拉伸/旋转/放样/扫描/平面/延伸/偏移/缝合/加厚，但**曲面裁剪（剪裁曲面）暂未
  封装**——`PreTrimSurface` / `PostTrimSurface` 的选择标记与参数组合还没试通，
  请先在界面里做，或用 `exec` 自行试验。也没有边界曲面（Boundary Surface）。
- 选面只支持 `all` 与 `near:x,y,z`（三角面片质心就近），没有按面索引/名称选择。
- 多实体零件：`volume_mm3` 已按实体求和，但 `mass_properties` 只报第一个实体。
- `close` 遇到未保存的改动会弹窗阻塞（无人值守脚本请先 `save`）。
- 只测过 Windows + SW2025 + 中文界面；英文界面按"第 1/2/3 个基准面"定位，理论可用但
  未实测。

## 安全

- 不联网、不下载、不写注册表，只调本机 COM + 读写你指定的文件。
- `exec` 会执行任意 Python（这是设计上的兜底口）。只运行你信任的输入。

## 目录

```
sw.cmd            Windows 快捷入口
sw.ps1            查找 Python 并调用 swbridge.py
swbridge.py       全部实现（COM 桥接 + 命令行）
docs/agent-skill.md   给 AI agent 的用法说明（可直接当技能文件）
docs/images/      示例截图
examples/         示例 STEP 与演示脚本
```

## License

[MIT](LICENSE) (c) 2026 matthewsu73

本项目与 Dassault Systèmes 无关。SOLIDWORKS 是 Dassault Systèmes 的商标，
使用本工具需要你自备合法授权的 SOLIDWORKS 副本。
