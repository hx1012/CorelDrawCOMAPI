# CorelDRAW X7 COM API 脚本参考

本仓库整理了通过阅读官方《CorelDRAW X7 脚本参考手册》（PDF / CHM）提炼的 VBA 宏使用指南、示例代码，以及 **C# 封装工具类**，涵盖形状对象、打印机操作、文字样式设置、表格创建与格式化等常用场景，方便开发者快速查阅正确的 COM API 用法。

---

## 仓库文件说明

| 文件 / 目录 | 说明 |
|-------------|------|
| `CorelStyleSheetHelper.cs` | **C# 工具类**：封装 StyleSheet / Style / StyleSet 的创建、应用、查找、导出等操作；**正确处理 `ss.Fill` 为 null 的问题** |
| `样式表操作工具类.md` | `CorelStyleSheetHelper.cs` 的完整中文文档：`ss.Fill` 为 null 的根本原因、两种修复路径、分场景示例、常见问题 |
| `TextRangeHelper.cs` | **C# 工具类**：统一从编辑选区/选中形状/当前页面/整个文档获取 `TextRange`，永不抛出 COM 异常 |
| `TextRange获取工具类.md` | `TextRangeHelper.cs` 的完整中文文档：API 说明、分场景示例、优先级说明、常见问题 |
| `Shape对象完整指南.md` | **Shape · ShapeRange · Shapes · SelectionInfo** 全部属性、方法、选择操作、布尔运算及综合示例 |
| `打印机操作完整指南.md` | **打印机系统 · PrintSettings · PostScript · 分色 · 陷印 · 印前处理**全部 API 速查及综合示例 |
| `对象样式设置指南.md` | 轮廓、填充、透明度、字符、段落、图文框、位图效果、QR 码等完整 API 速查 |
| `样式集与样式的区别.md` | Style 与 StyleSet 的概念对比、创建方法、正确 VBA 示例（含 `ss.Fill` 为 null 的修正） |
| `字体样式设置指南.md` | 字体名称、字号、粗/斜/下划线、大小写、上下标、颜色、间距等完整 API 速查 |
| `创建表格指南.md` | 表格创建、单元格读写、合并、填充、边框、行列管理等 API 速查及示例 |
| `CreateTable示例.vba` | 可运行的 VBA 宏（产品信息表 + 2009 年 1 月日历两个示例） |
| `CorelDRAW X7 脚本手册.chm` | CorelDRAW X7 官方脚本参考手册（CHM 格式） |
| `CorelDraw脚本参考手册X7_0*.pdf` | CorelDRAW X7 脚本参考手册分册（PDF 格式，共 8 册） |

---

## 运行环境

- **软件**：CorelDRAW X7 或更高版本（X8、2017–2021）
- **VBA**：无需安装额外插件，使用 CorelDRAW 内置宏引擎即可
- **C#**：.NET Framework 4.7.2 / .NET 6+，C# 8.0+；无需 COM 引用（`dynamic` 后期绑定）

---

## 快速开始

### VBA 宏

1. 打开 CorelDRAW。
2. 点击菜单 **工具（Tools）→ 宏（Macros）→ 宏编辑器（Macro Editor）**，或按 **Alt+F11**。
3. 在左侧"工程"窗格中插入新模块，将 `.vba` 文件中的代码粘贴进去。
4. 将光标置于任意 `Sub` 过程内，按 **F5** 运行。
5. 切换回 CorelDRAW 主窗口查看结果。

### C# 工具类（TextRangeHelper）

1. 将 [`TextRangeHelper.cs`](TextRangeHelper.cs) 复制到 C# 项目中。
2. 在代码文件顶部添加引用：
   ```csharp
   using System.Runtime.InteropServices;
   using CorelDrawCOMAPI;
   ```
3. 连接 CorelDRAW 并使用：
   ```csharp
   dynamic app = Marshal.GetActiveObject("CorelDRAW.Application");
   CorelTextRangeHelper.ForEachTextRange(app, tr =>
   {
       tr.Font = "微软雅黑";
       tr.Size = 12;
   });
   ```

---

## 内容导览

### StyleSheet 操作工具类（C#）→ [`CorelStyleSheetHelper.cs`](CorelStyleSheetHelper.cs) · [`样式表操作工具类.md`](样式表操作工具类.md)

封装了创建、应用、查找、导出样式和样式集的全套操作，并**正确处理了 `ss.Fill` 为 null 的问题**：

> **⚠️ 已知 API 陷阱**：`StyleSheet.CreateStyleSet()` 返回空容器，
> 其 `.Fill`、`.Outline` 等属性均为 `null`，直接赋值会报运行时错误。
> `Style`（样式集）对象**没有** `CreateStyle` 方法，唯一正确做法是改用 `CreateStyleSetFromShape`。

| 方法 | 说明 |
|------|------|
| `CreateFillStyle(app, name, r, g, b)` | 创建 RGB 纯色填充样式 |
| `CreateOutlineStyle(app, name, widthMm, r, g, b)` | 创建 RGB 轮廓样式 |
| `CreateCharacterStyle(app, name, font, size, ...)` | 创建字符样式 |
| `CreateStyleSetViaShape(...)` | **唯一正确方式**：通过临时形状创建样式集 |
| `ApplyStyle(shape, name)` / `ApplyStyleToShapes(shapes, name)` | 应用/批量应用样式 |
| `ExportStyles` / `ImportStyles` | 导出/导入 `.cdss` 样式文件 |

### TextRange 获取工具类（C#）→ [`TextRangeHelper.cs`](TextRangeHelper.cs) · [`TextRange获取工具类.md`](TextRange获取工具类.md)

统一封装了从四种来源获取 `TextRange` 的逻辑（不需要调用方关心当前是哪种状态）：

- **`TextRangeScope.EditingSelection`**：文字编辑模式下光标选中的区域（`Text.Selection`）
- **`TextRangeScope.SelectedShapes`**：选区中所有文字形状各自的完整文字（`Text.Story`）
- **`TextRangeScope.CurrentPage`**：当前页面所有文字形状的 Story
- **`TextRangeScope.CurrentDocument`**：整个文档所有页面所有文字形状的 Story
- **`TextRangeScope.Auto`**（默认）：按优先级自动检测，始终返回有意义的结果

核心方法：

| 方法 | 返回值 | 说明 |
|------|--------|------|
| `GetTextRanges(app, scope)` | `IReadOnlyList<dynamic>` | 永不返回 null；无文本时返回空列表 |
| `TryGetTextRange(app, scope)` | `dynamic`（可 null） | 取第一个 TextRange 或 null |
| `ForEachTextRange(app, action, scope)` | `void` | 批量操作，异常自动跳过 |

### Shape 对象 → [`Shape对象完整指南.md`](Shape对象完整指南.md)

涵盖以下四大类：

- **Shape**：CorelDRAW 中所有可见元素的核心对象
  - 位置与尺寸属性：`PositionX/Y`、`SizeWidth/Height`、`CenterX/Y`、`BoundingBox`
  - 变换方法：`Move`、`Rotate`、`RotateEx`、`SetSize`、`Stretch`、`Flip`、`Skew`、`AffineTransform`
  - 选择与层次：`Selected`、`AddToSelection`、`CreateSelection`、`OrderToFront`、`Layer`
  - 布尔运算：`Weld`、`Trim`、`Intersect`、`Combine`、`EqualDivide`
  - 克隆与复制：`Clone`、`Duplicate`、`StepAndRepeat`、`CloneLink`
  - 组合管理：`Group`、`Ungroup`、`UngroupAll`、`BreakApart`
  - 效果：`CreateBlend`、`CreateDropShadow`、`CreateContour`、`CreateEnvelope` 等 13 种
  - 形状访问：`Curve`、`Rectangle`、`Ellipse`、`Text`、`Bitmap`、`Polygon` 等

- **ShapeRange**：Shape 的动态数组，可用 `New ShapeRange` 创建
  - 集合管理：`Add`、`AddRange`、`Remove`、`RemoveAll`、`Sort`（CQL 排序）
  - 批量变换：与 Shape 相同的全套变换方法
  - 批量外观：`ApplyUniformFill`、`ApplyFountainFill`、`SetOutlineProperties` 等
  - 批量效果：`ApplyEffectBCI`、`ApplyEffectHSL` 等
  - 群组/合并：`Group`、`Combine`、`Ungroup`、`BreakApart`
  - 查询：`Exists`、`ExistsAnyOfType`、`CountAnyOfType`、`FindAnyOfType`

- **Shapes**：固定集合，反映文档真实结构
  - `All()`、`AllExcluding()`、`Range()`：转换为 ShapeRange
  - `FindShape()`、`FindShapes()`：按名称、类型、CQL 查询

- **SelectionInfo**：选区状态只读信息对象
  - `Can...` 系列属性（30+ 个）：判断当前选区能否执行指定操作
  - `Is...` 系列属性（40+ 个）：判断选区包含哪类对象
  - 关联形状访问：`BlendBottomShape`、`ContourControlShape`、`DropShadowGroup` 等

### 打印机操作 → [`打印机操作完整指南.md`](打印机操作完整指南.md)

涵盖以下模块：

- **SystemPrinters / Printer**：枚举已安装打印机、判断彩色/PostScript 能力
- **PrintSettings**（核心入口）：打印机选择、份数、范围、纸张、打印到文件；`Save`/`Load`/`ShowDialog` 方法
- **PrintOptions**：颜色模式（`PrnColorMode`）、位图/矢量/文字开关、降采样、栅格化打印
- **PrintLayout**：页面放置（`PrnPlaceType` 12 种）、平铺打印、出血
- **PrintPostScript**：PostScript Level 1/2/3、字体下载、JPEG 压缩、PDF 蒸馏选项
- **PrintPrepress**：裁切/套准标记、颜色校准条、密度计刻度、负片/镜像输出
- **PrintSeparations / SeparationPlate**：CMYK 分色、专色转换、加网角度/频率
- **PrintTrapping**：陷印宽度、黑色阈值、图像陷印模式（`PrnImageTrap`）
- **打印事件**：`BeforePrint`、`AfterPrint`、`QueryPrint`

### 对象样式设置 → [`对象样式设置指南.md`](对象样式设置指南.md)

涵盖以下样式类别：

- **轮廓**：`Shape.Outline`（线宽、颜色、线型、线端、箭头、合并模式）
- **填充**：`Shape.Fill`（均匀色、渐变、双色图样、纹理、PostScript）
- **透明度**：`Shape.Transparency`（均匀、渐变、图样、纹理、合并模式）
- **字符样式**：`Text.Story`（字体、字号、粗斜体、颜色、间距，详见字体指南）
- **段落样式**：缩进、行距、段前段后、制表位、对齐
- **图文框**：`Shape.TextFrame`（垂直对齐、分栏、链接、内边距）
- **位图效果**：转换为位图、PowerClip、透镜、阴影、轮廓图、调和
- **QR 码**：`CreateCustomShape("QRCode", ...)`（内容、纠错级别、颜色、Logo）

### 字体样式设置 → [`字体样式设置指南.md`](字体样式设置指南.md)

涵盖以下 `TextRange` 属性与方法：

- **字体 / 字号**：`Font`、`Size`
- **粗体 / 斜体**：`Bold`、`Italic`
- **字体样式枚举**：`Style`（`cdrFontStyle`，含 18 个常量）
- **下划线 / 删除线**：`Underline`、`Strikethru`（`cdrFontLine` 枚举）
- **大小写**：`Case`（`cdrFontCase`）、`ChangeCase`
- **上标 / 下标**：`Position`（`cdrFontPosition`）
- **文字颜色**：`Fill.ApplyUniformFill`
- **对齐方式**：`Alignment`
- **字符 / 词 / 行间距**：`CharSpacing`、`WordSpacing`、`SetLineSpacing`
- **字偶间距 / 旋转 / 偏移**：`RangeKerning`、`CharAngle`、`HorizShift`、`VertShift`

### 创建表格 → [`创建表格指南.md`](创建表格指南.md)

涵盖以下核心 API：

- **创建表格**：`CreateCustomShape("Table", ...)`（⚠️ 不存在 `CreateTable()` 方法）
- **访问 TableShape**：`Shape.Custom`
- **单元格访问**：`TableShape.Cell(Column, Row)`（列在前，行在后）
- **单元格文字**：`TableCell.TextShape.Text.Story`
- **填充 / 合并 / 边框 / 行列管理**

---

## 参考资料

- 本仓库中的《CorelDraw脚本参考手册X7》PDF（共 8 册）
- 本仓库中的《CorelDRAW X7 脚本手册》（CHM）
- CorelDRAW 官方开发者社区：https://community.coreldraw.com/sdk/

