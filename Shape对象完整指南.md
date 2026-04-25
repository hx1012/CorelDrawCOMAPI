# CorelDRAW X7 Shape / ShapeRange / Shapes / SelectionInfo 完整指南

> 内容来源：《CorelDRAW X7 脚本参考手册》第 1–5 册（本仓库 PDF）  
> 涵盖对象：**Shape · ShapeRange · Shapes · SelectionInfo**

---

## 目录

1. [对象关系总览](#一对象关系总览)
2. [Shape 对象](#二shape-对象)
   - 2.1 [Shape 属性一览](#21-shape-属性一览)
   - 2.2 [Shape 方法一览](#22-shape-方法一览)
   - 2.3 [位置与尺寸详解](#23-位置与尺寸详解)
   - 2.4 [变换操作详解](#24-变换操作详解)
   - 2.5 [选择与层次操作详解](#25-选择与层次操作详解)
   - 2.6 [布尔运算详解](#26-布尔运算详解)
   - 2.7 [克隆与复制详解](#27-克隆与复制详解)
   - 2.8 [组合与拆分详解](#28-组合与拆分详解)
3. [ShapeRange 对象](#三shaperange-对象)
   - 3.1 [ShapeRange 属性一览](#31-shaperange-属性一览)
   - 3.2 [ShapeRange 方法一览](#32-shaperange-方法一览)
   - 3.3 [创建与管理 ShapeRange](#33-创建与管理-shaperange)
   - 3.4 [批量操作详解](#34-批量操作详解)
4. [Shapes 集合](#四shapes-集合)
   - 4.1 [Shapes 属性与方法一览](#41-shapes-属性与方法一览)
   - 4.2 [查找与筛选详解](#42-查找与筛选详解)
5. [SelectionInfo 对象](#五selectioninfo-对象)
   - 5.1 [SelectionInfo 属性一览](#51-selectioninfo-属性一览)
   - 5.2 [实用示例](#52-实用示例)
6. [选择操作（Selection API）](#六选择操作selection-api)
7. [cdrShapeType 枚举常量](#七cdrshapetype-枚举常量)
8. [综合实战示例](#八综合实战示例)

---

## 一、对象关系总览

```
Application
├── ActiveShape          → Shape（当前激活的单个形状）
├── ActiveSelection      → Shape（类型为 cdrSelectionShape 的选区对象）
└── ActiveSelectionRange → ShapeRange（当前选区的 ShapeRange 表示）

Document
├── Selection()          → Shape（选区对象，与 ActiveSelection 等价）
├── SelectionRange       → ShapeRange（选区的 ShapeRange 表示）
└── SelectionInfo        → SelectionInfo（选区状态信息）

Page / Layer / Group-Shape
└── Shapes               → Shapes（该容器中所有形状的集合）
    ├── .All()           → ShapeRange（全部形状的范围）
    ├── .FindShape(...)  → Shape（单个形状）
    ├── .FindShapes(...) → ShapeRange（多个形状）
    ├── .Range(...)      → ShapeRange（指定索引的范围）
    └── .Item(n)         → Shape（按索引或名称取单个形状）

ShapeRange                         （动态数组，可 New 创建，可手工增删）
└── .Item(n) / sr(n)   → Shape

Shape
└── .Shapes            → Shapes（仅对 GroupShape 有效，子形状集合）
```

**常用全局快捷入口**（无需写 `ActiveDocument.`）：

| 快捷写法 | 等价完整路径 | 说明 |
|---------|------------|------|
| `ActiveShape` | `Application.ActiveShape` | 当前被选中的单个形状 |
| `ActiveSelection` | `Application.ActiveSelection` | 选区 Shape 对象 |
| `ActiveSelectionRange` | `Application.ActiveSelectionRange` | 选区 ShapeRange |
| `ActivePage.Shapes` | `ActiveDocument.ActivePage.Shapes` | 当前页面的形状集合 |
| `ActiveLayer.Shapes` | `ActiveDocument.ActiveLayer.Shapes` | 当前图层的形状集合 |

---

## 二、Shape 对象

**描述**：`Shape` 类代表 CorelDRAW 中的一个形状对象（矩形、椭圆、曲线、文字、群组、选区等）。所有在画面上的可见元素，包括选区本身，都以 `Shape` 对象的形式暴露在 API 中。

### 2.1 Shape 属性一览

#### 基本信息属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Type` | `cdrShapeType` | 只读 | 形状类型（矩形/椭圆/曲线/文字/组合等，见第七节）|
| `Name` | `String` | 读写 | 形状名称（未设置时返回空字符串）|
| `StaticID` | `Long` | 只读 | 形状的唯一静态 ID（文档内唯一，保存后仍保持）|
| `ZOrder` | `Long` | 只读 | 形状在父容器集合中的索引位置（Z 轴顺序）|
| `Application` | `Application` | 只读 | 顶层应用对象 |
| `Parent` | `Object` | 只读 | 父对象（图层或组合）|
| `Layer` | `Layer` | 读写 | 所在图层（赋值可将形状移至指定图层）|
| `Page` | `Page` | 只读 | 所在页面 |
| `ParentGroup` | `Shape` | 只读 | 直接父组合形状（若未在组中则为 `Nothing`）|
| `IsSimpleShape` | `Boolean` | 只读 | 是否为简单形状（非组合、非效果组）|
| `Virtual` | `Boolean` | 只读 | 是否为虚拟形状（控制手柄等不可打印的辅助对象）|
| `Properties` | `DataItems` | 只读 | 附加到形状上的自定义数据集合 |
| `ObjectData` | `DataItems` | 只读 | 对象数据（与 `Properties` 类似）|

#### 位置与尺寸属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `PositionX` | `Double` | 读写 | 形状的水平坐标（相对于 `Document.ReferencePoint`）|
| `PositionY` | `Double` | 读写 | 形状的垂直坐标 |
| `CenterX` | `Double` | 读写 | 形状中心的水平坐标 |
| `CenterY` | `Double` | 读写 | 形状中心的垂直坐标 |
| `SizeWidth` | `Double` | 读写 | 形状宽度（赋值会拉伸形状）|
| `SizeHeight` | `Double` | 读写 | 形状高度 |
| `LeftX` | `Double` | 只读 | 形状左边界 X 坐标 |
| `RightX` | `Double` | 只读 | 形状右边界 X 坐标 |
| `TopY` | `Double` | 只读 | 形状上边界 Y 坐标 |
| `BottomY` | `Double` | 只读 | 形状下边界 Y 坐标 |
| `BoundingBox` | `Rect` | 只读 | 包围盒矩形对象 |
| `OriginalWidth` | `Double` | 只读 | 创建时的原始宽度 |
| `OriginalHeight` | `Double` | 只读 | 创建时的原始高度 |

#### 变换属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `RotationAngle` | `Double` | 读写 | 旋转角度（赋值会将形状旋转到指定角度）|
| `RotationCenterX` | `Double` | 读写 | 旋转中心 X 坐标 |
| `RotationCenterY` | `Double` | 读写 | 旋转中心 Y 坐标 |
| `AbsoluteHScale` | `Double` | 只读 | 从创建到现在累计的水平缩放比例 |
| `AbsoluteVScale` | `Double` | 只读 | 从创建到现在累计的垂直缩放比例 |
| `AbsoluteSkew` | `Double` | 只读 | 从创建到现在累计的倾斜角度 |

#### 外观属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Fill` | `Fill` | 只读* | 填充属性对象（*可通过子方法修改）|
| `Outline` | `Outline` | 只读* | 轮廓属性对象 |
| `Transparency` | `Transparency` | 只读* | 透明度属性对象 |
| `CanHaveFill` | `Boolean` | 只读 | 是否支持填充 |
| `CanHaveOutline` | `Boolean` | 只读 | 是否支持轮廓 |
| `FillMode` | `cdrFillMode` | 读写 | 填充模式（奇偶规则/非零规则）|
| `DrapeFill` | `Boolean` | 读写 | 填充是否随形变换 |
| `OverprintFill` | `Boolean` | 读写 | 填充是否叠印 |
| `OverprintOutline` | `Boolean` | 读写 | 轮廓是否叠印 |
| `OverprintBitmap` | `Boolean` | 读写 | 位图是否叠印 |
| `PixelAlignedRendering` | `Boolean` | 读写 | 是否启用像素对齐渲染 |
| `Spread` | `Double` | 只读 | 扩散量（用于陷印）|
| `WrapText` | `cdrWrapStyle` | 读写 | 段落文字绕排方式 |
| `TextWrapOffset` | `Double` | 读写 | 文字绕排偏移量 |
| `URL` | `String` | 读写 | 超链接 URL |

#### 选择与锁定属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Selected` | `Boolean` | 读写 | 是否被选中（赋值可加入/移出选区）|
| `Selectable` | `Boolean` | 只读 | 是否可被选中（可见且未锁定）|
| `Locked` | `Boolean` | 读写 | 是否被锁定 |

#### 导航属性

| 属性 | 类型 | 说明 |
|------|------|------|
| `Next([Level, EnterGroups, Loop])` | `Shape` | 返回下一个形状（可按页/图层/组范围导航）|
| `Previous([Level, EnterGroups, Loop])` | `Shape` | 返回上一个形状 |
| `Shapes` | `Shapes` | 子形状集合（仅对 GroupShape 有效）|
| `SnapPoints` | `SnapPoints` | 吸附点集合 |
| `TreeNode` | `TreeNode` | 对象管理器中的树节点 |

#### 形状特有属性（按 Type 使用）

| 属性 | 类型 | 说明 |
|------|------|------|
| `Curve` | `Curve` | 曲线数据（`cdrCurveShape`）|
| `DisplayCurve` | `Curve` | 显示曲线（所有类型均有）|
| `Rectangle` | `Rectangle` | 矩形参数（`cdrRectangleShape`）|
| `Ellipse` | `Ellipse` | 椭圆参数（`cdrEllipseShape`）|
| `Polygon` | `Polygon` | 多边形参数（`cdrPolygonShape`）|
| `Text` | `Text` | 文字对象（`cdrTextShape`）|
| `Bitmap` | `Bitmap` | 位图对象（`cdrBitmapShape`）|
| `BSpline` | `BSpline` | B 样条对象 |
| `OLE` | `OLE` | OLE 对象（`cdrOLEObjectShape`）|
| `EPS` | `EPS` | EPS 对象 |
| `Connector` | `Connector` | 连接线对象（`cdrConnectorShape`）|
| `Guide` | `Guide` | 辅助线对象（`cdrGuidelineShape`）|
| `Symbol` | `Symbol` | 符号对象（`cdrSymbolShape`）|
| `Custom` | `Custom` | 自定义对象（`cdrCustomShape`，如 Table）|
| `Dimension` | `Dimension` | 标注对象 |
| `PowerClip` | `PowerClip` | PowerClip 容器属性 |
| `PowerClipParent` | `PowerClip` | 若本身在 PowerClip 中，返回容器 |
| `Effect` | `Effect` | 效果对象（用于效果控制形状）|
| `Effects` | `Effects` | 所有效果集合 |
| `CloneLink` | `CloneLink` | 克隆链接属性 |
| `Clones` | `ShapeRange` | 本控制形状的所有克隆 |
| `Style` | `Style` | 应用的样式 |

---

### 2.2 Shape 方法一览

#### 位置与变换

| 方法 | 签名 | 说明 |
|------|------|------|
| `Move` | `(DeltaX, DeltaY)` | 按偏移量移动形状 |
| `SetPosition` | `(X, Y)` | 移动到绝对坐标（相对于 ReferencePoint）|
| `SetPositionEx` | `(ReferencePoint, X, Y)` | 移动到绝对坐标，指定参考点 |
| `GetPosition` | `(X, Y)` | 获取当前坐标（输出参数）|
| `GetPositionEx` | `(ReferencePoint, X, Y)` | 获取坐标，指定参考点 |
| `SetSize` | `([Width], [Height])` | 设置尺寸（省略其一则等比缩放）|
| `SetSizeEx` | `(CenterX, CenterY, [Width], [Height])` | 以指定锚点为中心设置尺寸 |
| `GetSize` | `(Width, Height)` | 获取当前尺寸（输出参数）|
| `SetBoundingBox` | `(X, Y, Width, Height, ...)` | 设置包围盒（含多种对齐选项）|
| `GetBoundingBox` | `(X, Y, Width, Height, [UseOutline])` | 获取包围盒（输出参数）|
| `Rotate` | `(Angle)` | 绕当前旋转中心旋转 |
| `RotateEx` | `(Angle, CenterX, CenterY)` | 绕指定点旋转 |
| `SetRotationCenter` | `(X, Y)` | 设置旋转中心 |
| `Skew` | `(AngleX, AngleY)` | 倾斜 |
| `SkewEx` | `(CenterX, CenterY, AngleX, AngleY)` | 绕指定点倾斜 |
| `Stretch` | `(StretchX, [StretchY], [StretchChars])` | 按比例拉伸 |
| `StretchEx` | `(CenterX, CenterY, StretchX, [StretchY])` | 绕指定点拉伸 |
| `Flip` | `(Axes)` | 水平/垂直翻转（`cdrFlipAxes`）|
| `AffineTransform` | `(a, b, c, d, e, f)` | 应用仿射变换矩阵 |
| `TransformMatrix` | `(Matrix)` | 应用变换矩阵 |
| `GetMatrix` | `(a, b, c, d, e, f)` | 获取当前变换矩阵 |
| `SetMatrix` | `(a, b, c, d, e, f)` | 设置变换矩阵 |
| `ClearTransformations` | `()` | 清除所有变换（还原到创建时状态）|
| `Project` | `(...)` | 透视变形投影 |
| `Unproject` | `()` | 移除透视投影 |

#### 编辑操作

| 方法 | 签名 | 说明 |
|------|------|------|
| `Copy` | `()` | 复制到剪贴板 |
| `Cut` | `()` | 剪切到剪贴板 |
| `Delete` | `()` | 删除形状 |
| `Duplicate` | `([OffsetX], [OffsetY]) As Shape` | 就地复制（返回新形状）|
| `DuplicateAsRange` | `([OffsetX], [OffsetY]) As ShapeRange` | 复制并以 ShapeRange 返回 |
| `Clone` | `([OffsetX], [OffsetY]) As Shape` | 克隆（与原形状保持链接）|
| `CloneAsRange` | `([OffsetX], [OffsetY]) As ShapeRange` | 克隆并以 ShapeRange 返回 |
| `StepAndRepeat` | `(NumCopies, DistanceX, DistanceY, ...) As ShapeRange` | 步骤与重复（批量复制并排列）|
| `CopyPropertiesFrom` | `(Source, Properties) As Boolean` | 从另一形状复制属性 |
| `ReplaceWith` | `(Shape)` | 用另一形状替换本形状 |

#### 布尔运算

| 方法 | 签名 | 说明 |
|------|------|------|
| `Combine` | `() As Shape` | 合并（适用于选区或组合形状）|
| `Weld` | `(TargetShape, ...) As Shape` | 焊接（合并两个形状的外轮廓）|
| `Trim` | `(TargetShape, [LeaveSource], [LeaveTarget]) As Shape` | 修剪 |
| `Intersect` | `(TargetShape, [LeaveSource], [LeaveTarget]) As Shape` | 取交集 |
| `EqualDivide` | `(NumDivisions, ...) As ShapeRange` | 等分切割 |

#### 形状转换

| 方法 | 签名 | 说明 |
|------|------|------|
| `ConvertToCurves` | `()` | 转换为曲线 |
| `ConvertToBitmap` | `([BitDepth], [Grayscale], ...) As Shape` | 转换为位图（旧版，建议用 Ex）|
| `ConvertToBitmapEx` | `([Mode], [Dithered], [Transparent], ...) As Shape` | 转换为位图（推荐）|
| `ConvertToSymbol` | `(Name) As Shape` | 转换为符号 |
| `CreateArrowHead` | `() As ArrowHead` | 从形状创建箭头 |

#### 群组管理

| 方法 | 签名 | 说明 |
|------|------|------|
| `Group` | `() As Shape` | 将选区中的形状组合（仅对 SelectionShape 有效）|
| `Ungroup` | `()` | 取消一层组合 |
| `UngroupAll` | `()` | 递归取消所有嵌套组合 |
| `UngroupAllEx` | `() As ShapeRange` | 递归取消组合并返回结果 ShapeRange |
| `UngroupEx` | `() As ShapeRange` | 取消一层组合并返回结果 ShapeRange |
| `BreakApart` | `()` | 拆分曲线的子路径，或拆分文字为单词/字符 |
| `BreakApartEx` | `() As ShapeRange` | 拆分并返回 ShapeRange |

#### 选择与层次

| 方法 | 签名 | 说明 |
|------|------|------|
| `CreateSelection` | `()` | 创建仅包含本形状的选区 |
| `AddToSelection` | `()` | 加入当前选区 |
| `RemoveFromSelection` | `()` | 从当前选区中移除 |
| `MoveToLayer` | `(Layer)` | 移到指定图层 |
| `CopyToLayer` | `(Layer)` | 复制到指定图层 |
| `CopyToLayerAsRange` | `(Layer) As ShapeRange` | 复制到图层并返回 ShapeRange |
| `OrderToFront` | `()` | 移到最前 |
| `OrderToBack` | `()` | 移到最后 |
| `OrderForwardOne` | `()` | 向前移一层 |
| `OrderBackOne` | `()` | 向后移一层 |
| `OrderFrontOf` | `(Shape)` | 移到指定形状前面 |
| `OrderBackOf` | `(Shape)` | 移到指定形状后面 |
| `OrderIsInFrontOf` | `(Shape) As Boolean` | 判断是否在另一形状前面 |
| `OrderReverse` | `()` | 反转组/选区中所有形状的堆叠顺序 |

#### 效果操作

| 方法 | 签名 | 说明 |
|------|------|------|
| `CreateBlend` | `(TargetShape, ...) As Effect` | 创建调和效果 |
| `CreateBoundary` | `(X, Y, ...) As Shape` | 创建边界曲线 |
| `CreateContour` | `(Type, Offset, ...) As Effect` | 创建轮廓图效果 |
| `CreateDropShadow` | `(OffsetX, OffsetY, ...) As Effect` | 创建阴影效果 |
| `CreateEnvelope` | `(...) As Effect` | 创建封套效果 |
| `CreateEnvelopeFromCurve` | `(Curve, ...) As Effect` | 从曲线创建封套 |
| `CreateEnvelopeFromShape` | `(Shape, ...) As Effect` | 从形状创建封套 |
| `CreateExtrude` | `(...) As Effect` | 创建立体效果 |
| `CreateLens` | `(Type, ...) As Effect` | 创建透镜效果 |
| `CreatePerspective` | `(...) As Effect` | 创建透视效果 |
| `CreatePushPullDistortion` | `(X, Y, ...) As Effect` | 创建推拉变形 |
| `CreateCustomDistortion` | `(Curve)` | 从曲线创建自定义变形 |
| `CreateTwisterDistortion` | `(X, Y, ...) As Effect` | 创建扭曲变形 |
| `CreateZipperDistortion` | `(X, Y, ...) As Effect` | 创建拉链变形 |
| `ClearEffect` | `(EffectType)` | 清除指定类型的效果 |
| `Separate` | `()` | 拆分效果组（分离控制形状和效果形状）|
| `ApplyStyle` | `(StyleName, [StyleSet])` | 应用样式 |
| `ApplyEffectBCI` | `(Brightness, Contrast, Intensity)` | 亮度/对比度/强度效果 |
| `ApplyEffectColorBalance` | `(CyanRed, MagentaGreen, YellowBlue, ...)` | 色彩平衡效果 |
| `ApplyEffectGamma` | `(Gamma)` | 伽马校正效果 |
| `ApplyEffectHSL` | `(Hue, Saturation, Lightness)` | 色相/饱和度/亮度效果 |
| `ApplyEffectInvert` | `()` | 反色效果 |
| `ApplyEffectPosterize` | `(Levels)` | 色调分离效果 |

#### 其他方法

| 方法 | 签名 | 说明 |
|------|------|------|
| `AddToPowerClip` | `(Container, [CenterInContainer])` | 放入 PowerClip 容器 |
| `RemoveFromContainer` | `([Level])` | 从 PowerClip 容器中取出 |
| `PlaceTextInside` | `(Shape)` | 将文字置入形状内部 |
| `Fillet` | `(Radius, [CombineSmoothSegments])` | 圆角 |
| `Chamfer` | `(DistanceA, DistanceB, [CombineSmoothSegments])` | 倒角 |
| `Scallop` | `(Radius, [CombineSmoothSegments])` | 扇形角 |
| `FindSnapPoint` | `(X, Y, ...) As SnapPoint` | 查找最近的吸附点 |
| `SnapPointsOfType` | `(Type) As SnapPoints` | 返回指定类型的吸附点集合 |
| `GetLinkedShapes` | `() As ShapeRange` | 获取链接的形状 |
| `GetOverprintFillState` | `() As cdrTriState` | 获取填充叠印状态 |
| `GetOverprintOutlineState` | `() As cdrTriState` | 获取轮廓叠印状态 |
| `IsOnShape` | `(X, Y, ...) As Long` | 判断坐标是否在形状上 |
| `IsTypeAnyOf` | `(ParamArray TypeList)` | 判断形状类型是否为指定类型之一 |
| `CompareTo` | `(Shape2, CompareType, [Condition]) As Boolean` | 与另一形状比较属性 |
| `CompareToEx` | `(Shape2, Condition) As Boolean` | 用 CQL 表达式比较形状 |
| `Evaluate` | `(Expression) As Variant` | 用 CQL 对形状求值 |
| `AlignToGrid` | `()` | 对齐到网格 |
| `AlignToPage` | `(Alignment)` | 对齐到页面边缘 |
| `AlignToPageCenter` | `(Alignment)` | 对齐到页面中心 |
| `AlignToPoint` | `(Alignment, X, Y)` | 对齐到指定点 |
| `AlignToShape` | `(Alignment, Shape)` | 对齐到另一形状 |
| `AlignToShapeRange` | `(Alignment, ShapeRange)` | 对齐到形状范围 |
| `AlignAndDistribute` | `(AlignH, AlignV, DistH, DistV, ...)` | 对齐与分布 |
| `Distribute` | `(Type, ...)` | 分布形状 |
| `CreateDocumentFrom` | `() As Document` | 从选区创建新文档 |
| `CustomCommand` | `(CommandName, ...)` | 执行自定义命令 |
| `RestoreCloneLink` | `(...)` | 恢复克隆链接 |

---

### 2.3 位置与尺寸详解

```vba
' 获取形状的位置和尺寸
Sub GetShapeInfo()
    Dim s As Shape
    Dim x As Double, y As Double, w As Double, h As Double
    Set s = ActiveShape
    ' 方法 1：直接属性
    Debug.Print "中心：(" & s.CenterX & ", " & s.CenterY & ")"
    Debug.Print "宽高：" & s.SizeWidth & " × " & s.SizeHeight
    Debug.Print "边界：L=" & s.LeftX & " R=" & s.RightX & " T=" & s.TopY & " B=" & s.BottomY
    ' 方法 2：GetBoundingBox（更精确，可含轮廓宽度）
    s.GetBoundingBox x, y, w, h, True   ' True = 含轮廓宽度
    Debug.Print "包围盒左下：(" & x & ", " & y & ")  尺寸：" & w & " × " & h
End Sub

' 移动形状到页面中心
Sub CenterOnPage()
    Dim s As Shape
    Set s = ActiveShape
    s.SetPositionEx cdrCenter, ActivePage.SizeWidth / 2, ActivePage.SizeHeight / 2
End Sub

' 将所有形状调整到同一尺寸（1英寸正方形）
Sub ResizeAll()
    Dim s As Shape
    ActiveDocument.ReferencePoint = cdrCenter
    For Each s In ActivePage.Shapes
        If s.Type <> cdrGuidelineShape Then
            s.SetSize 1, 1
        End If
    Next s
End Sub

' 强制所有形状在页面范围内
Sub ClampToPage()
    Dim s As Shape
    Dim pw As Double, ph As Double
    ActivePage.GetSize pw, ph
    For Each s In ActivePage.Shapes
        If s.Type <> cdrGuidelineShape Then
            ActiveDocument.ReferencePoint = cdrBottomLeft
            If s.PositionX < 0 Then s.PositionX = 0
            If s.PositionY < 0 Then s.PositionY = 0
            ActiveDocument.ReferencePoint = cdrTopRight
            If s.PositionX > pw Then s.PositionX = pw
            If s.PositionY > ph Then s.PositionY = ph
        End If
    Next s
End Sub
```

---

### 2.4 变换操作详解

```vba
' 旋转：绕形状当前旋转中心旋转 45°
Sub RotateShape()
    ActiveShape.Rotate 45
End Sub

' 旋转：绕页面中心旋转，创建花朵效果
Sub FlowerEffect()
    Const NUM As Long = 12
    Dim s As Shape
    Dim i As Long
    Set s = ActiveLayer.CreateEllipse(5, 5, 7, 6)
    s.RotationCenterX = 4
    s.RotationCenterY = 5
    s.Fill.UniformColor.RGBAssign 255, 100, 100
    For i = 1 To NUM - 1
        s.Duplicate.RotateEx 360 / NUM * i, 5, 5
    Next i
End Sub

' 翻转：水平镜像
Sub MirrorHorizontal()
    ActiveShape.Flip cdrFlipHorizontal
End Sub

' 倾斜
Sub SkewShape()
    ActiveShape.Skew 15, 0   ' 水平倾斜 15°
End Sub

' 按比例拉伸到 200%
Sub DoubleSize()
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveShape.Stretch 2
End Sub

' 清除所有变换
Sub ResetTransform()
    ActiveShape.ClearTransformations
End Sub

' 步骤与重复：创建 5×3 网格
Sub StepRepeat()
    Dim s As Shape
    Dim sr As ShapeRange
    Set s = ActiveLayer.CreateRectangle2(0, 0, 1, 1)
    s.Fill.UniformColor.RGBAssign 255, 200, 100
    ' 先水平复制 4 份（间距 1.5）
    Set sr = s.StepAndRepeat(4, 1.5, 0)
    sr.AddRange sr.StepAndRepeat(2, 0, -1.5)
    sr.AddRange s.StepAndRepeat(2, 0, -1.5)
End Sub
```

---

### 2.5 选择与层次操作详解

```vba
' 将形状添加到当前选区
Sub SelectAdditionally()
    Dim s As Shape
    Set s = ActivePage.FindShape(Name:="MyRect")
    If Not s Is Nothing Then s.AddToSelection
End Sub

' 从选区中移除所有文字形状
Sub DeselectText()
    Dim s As Shape
    For Each s In ActiveSelection.Shapes
        If s.Type = cdrTextShape Then s.Selected = False
    Next s
End Sub

' 将选中形状移到顶层
Sub BringToFront()
    Dim s As Shape
    For Each s In ActiveSelectionRange
        s.OrderToFront
    Next s
End Sub

' 将所有矩形移到新图层
Sub MoveRectsToLayer()
    Dim s As Shape
    Dim lr As Layer
    Set lr = ActivePage.CreateLayer("矩形图层")
    For Each s In ActivePage.FindShapes(Type:=cdrRectangleShape)
        s.Layer = lr
    Next s
End Sub

' 检查堆叠顺序
Sub CheckOrder()
    Dim s1 As Shape, s2 As Shape
    Set s1 = ActivePage.Shapes.First
    Set s2 = ActivePage.Shapes.Last
    If s1.OrderIsInFrontOf(s2) Then
        MsgBox "第一个形状在第二个形状前面（Z轴更高）"
    End If
End Sub
```

---

### 2.6 布尔运算详解

```vba
' 焊接两个形状
Sub WeldShapes()
    Dim s1 As Shape, s2 As Shape
    Set s1 = ActiveLayer.CreateRectangle2(0, 0, 3, 3)
    Set s2 = ActiveLayer.CreateEllipse2(2, 2, 1.5)
    Set s1 = s1.Weld(s2)   ' s1 成为焊接后的新形状
    s1.Fill.UniformColor.RGBAssign 0, 120, 255
End Sub

' 修剪：创建月牙形
Sub CreateCrescent()
    Dim s1 As Shape, s2 As Shape
    Set s1 = ActiveLayer.CreateEllipse2(3, 6, 2)
    Set s2 = ActiveLayer.CreateEllipse2(4, 6, 2)
    s2.Trim s1         ' 用 s2 修剪 s1
    s2.Delete
End Sub

' 取交集：RGB 色彩叠加示意图
Sub RGBDiagram()
    Dim circles(2) As Shape
    Dim i As Long
    Dim cx As Double, cy As Double
    cx = 5: cy = 5
    For i = 0 To 2
        Set circles(i) = ActiveLayer.CreateEllipse2( _
            cx + 1.2 * Cos(i * 2.094), _
            cy + 1.2 * Sin(i * 2.094), 1.5)
    Next i
    circles(0).Fill.UniformColor.RGBAssign 255, 0, 0
    circles(1).Fill.UniformColor.RGBAssign 0, 255, 0
    circles(2).Fill.UniformColor.RGBAssign 0, 0, 255
    ' 两两取交集
    Dim s01 As Shape, s12 As Shape, s02 As Shape, s012 As Shape
    Set s01 = circles(0).Intersect(circles(1))
    s01.Fill.UniformColor.RGBAssign 255, 255, 0   ' 红+绿=黄
    Set s12 = circles(1).Intersect(circles(2))
    s12.Fill.UniformColor.RGBAssign 0, 255, 255   ' 绿+蓝=青
    Set s02 = circles(0).Intersect(circles(2))
    s02.Fill.UniformColor.RGBAssign 255, 0, 255   ' 红+蓝=品红
    ' 三者交集
    Set s012 = s01.Intersect(s12)
    s012.Fill.UniformColor.RGBAssign 255, 255, 255 ' 三色=白
End Sub
```

---

### 2.7 克隆与复制详解

```vba
' 克隆：修改克隆的填充不影响控制形状
Sub CloneDemo()
    Dim ctrl As Shape, cln As Shape
    Set ctrl = ActiveLayer.CreateRectangle2(0, 0, 3, 3)
    ctrl.Fill.UniformColor.RGBAssign 255, 0, 0
    Set cln = ctrl.Clone(4, 0)
    cln.Fill.UniformColor.RGBAssign 0, 255, 0   ' 克隆使用新颜色
End Sub

' 通过克隆链接访问控制形状
Sub AccessCloneParent()
    Dim s As Shape
    Set s = ActiveShape   ' 假设选中了一个克隆
    If Not s.CloneLink Is Nothing Then
        ' 修改控制形状的填充
        s.CloneLink.CloneParent.Fill.ApplyTextureFill "Stone", "Samples"
    End If
End Sub

' 批量克隆并排列（环形）
Sub CircularClones()
    Const N As Long = 8
    Dim ctrl As Shape
    Dim i As Long
    Set ctrl = ActiveLayer.CreateEllipse2(5, 7, 0.3)
    ctrl.Fill.UniformColor.RGBAssign 255, 100, 0
    For i = 1 To N - 1
        Dim cln As Shape
        Set cln = ctrl.Clone(0, 0)
        cln.SetPositionEx cdrCenter, _
            5 + 2 * Cos(i * 2 * 3.14159 / N), _
            5 + 2 * Sin(i * 2 * 3.14159 / N)
    Next i
End Sub
```

---

### 2.8 组合与拆分详解

```vba
' 组合选区中的形状
Sub GroupSelected()
    Dim grp As Shape
    If ActiveSelection.Shapes.Count < 2 Then
        MsgBox "请先选中至少两个形状"
        Exit Sub
    End If
    Set grp = ActiveSelection.Group
    grp.Name = "我的组合"
End Sub

' 合并（Combine）：创建镂空效果
Sub CombineHollow()
    Dim outer As Shape, inner As Shape
    Set outer = ActiveLayer.CreateEllipse2(5, 5, 3)
    Set inner = ActiveLayer.CreateEllipse2(5, 5, 1.5)
    outer.AddToSelection
    inner.AddToSelection
    Dim combined As Shape
    Set combined = ActiveSelection.Combine
    combined.Fill.UniformColor.RGBAssign 0, 120, 200
End Sub

' 遍历组合内部的形状
Sub TraverseGroup()
    Dim s As Shape
    For Each s In ActivePage.Shapes
        If s.Type = cdrGroupShape Then
            Dim child As Shape
            For Each child In s.Shapes
                Debug.Print "  子形状：" & child.Type
            Next child
        End If
    Next s
End Sub

' 递归遍历所有形状（含嵌套组）
Sub TraverseAllShapes()
    TraverseShapes ActivePage.Shapes
End Sub
Private Sub TraverseShapes(ss As Shapes)
    Dim s As Shape
    For Each s In ss
        If s.Type = cdrGroupShape Then
            TraverseShapes s.Shapes
        Else
            Debug.Print "形状：" & s.Name & "  类型：" & s.Type
        End If
    Next s
End Sub
```

---

## 三、ShapeRange 对象

**描述**：`ShapeRange` 是 `Shape` 对象的**动态数组**。可以用 `New ShapeRange` 手工创建，也可从各类查找/枚举方法中获得。对 `ShapeRange` 的大多数操作（移动、旋转、填充等）会作用于其中的**每一个形状**。

### 3.1 ShapeRange 属性一览

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Count` | `Long` | 只读 | 范围中的形状数量 |
| `Item(IndexOrName)` | `Shape` | 只读 | 按索引（从 1 开始）或名称取形状（默认属性，可省略）|
| `FirstShape` | `Shape` | 只读 | 第一个形状 |
| `LastShape` | `Shape` | 只读 | 最后一个形状 |
| `PositionX` | `Double` | 读写 | 整个范围的参考点 X 坐标 |
| `PositionY` | `Double` | 读写 | 整个范围的参考点 Y 坐标 |
| `CenterX` | `Double` | 只读 | 包围盒中心 X |
| `CenterY` | `Double` | 只读 | 包围盒中心 Y |
| `SizeWidth` | `Double` | 只读 | 包围盒宽度 |
| `SizeHeight` | `Double` | 只读 | 包围盒高度 |
| `LeftX` | `Double` | 只读 | 包围盒左边界 X |
| `RightX` | `Double` | 只读 | 包围盒右边界 X |
| `TopY` | `Double` | 只读 | 包围盒上边界 Y |
| `BottomY` | `Double` | 只读 | 包围盒下边界 Y |
| `BoundingBox` | `Rect` | 只读 | 包围盒 Rect 对象 |
| `RotationCenterX` | `Double` | 读写 | 旋转中心 X |
| `RotationCenterY` | `Double` | 读写 | 旋转中心 Y |
| `Shapes` | `Shapes` | 只读 | 范围中所有形状的 Shapes 集合 |
| `Parent` | `Object` | 只读 | 父对象 |
| `Application` | `Application` | 只读 | 顶层应用对象 |
| `ReverseRange` | `ShapeRange` | 只读 | 反序后的范围 |

---

### 3.2 ShapeRange 方法一览

#### 集合管理

| 方法 | 签名 | 说明 |
|------|------|------|
| `Add` | `(Shape)` | 添加单个形状 |
| `AddRange` | `(ShapeRange)` | 添加另一个范围中的所有形状 |
| `Remove` | `(Index)` | 按索引从范围中移除（不删除文档中的形状）|
| `RemoveAll` | `()` | 清空范围（不删除文档中的形状）|
| `RemoveRange` | `(ShapeRange)` | 从范围中移除另一个范围中的形状 |
| `DeleteItem` | `(Index)` | 按索引从范围中移除并从文档中删除形状 |
| `IndexOf` | `(Shape) As Long` | 返回形状在范围中的索引（0 = 不存在）|
| `Exists` | `(Shape) As Boolean` | 判断形状是否在范围中 |
| `ExistsAnyOfType` | `(TypeList...)` | 判断范围中是否有指定类型的形状 |
| `CountAnyOfType` | `(TypeList...) As Long` | 统计指定类型形状的数量 |
| `FindAnyOfType` | `(TypeList...) As ShapeRange` | 返回指定类型形状的子范围 |
| `All` | `() As ShapeRange` | 返回自身的副本 |
| `AllExcluding` | `(IndexArray) As ShapeRange` | 返回排除指定索引后的副本 |
| `Range` | `(IndexArray) As ShapeRange` | 返回指定索引子集 |
| `Sort` | `(Expression, [Start], [End], [Data])` | 按 CQL 表达式排序 |

#### 位置与变换

| 方法 | 说明 |
|------|------|
| `Move(DeltaX, DeltaY)` | 整体移动 |
| `SetPosition(X, Y)` | 移动到绝对位置 |
| `SetPositionEx(RefPoint, X, Y)` | 指定参考点移动 |
| `GetPosition(X, Y)` | 获取当前位置 |
| `GetPositionEx(RefPoint, X, Y)` | 获取当前位置（指定参考点）|
| `SetSize([W], [H])` | 调整尺寸 |
| `SetSizeEx(CX, CY, [W], [H])` | 以指定锚点调整尺寸 |
| `GetSize(W, H)` | 获取尺寸 |
| `SetBoundingBox(X, Y, W, H, ...)` | 设置包围盒 |
| `GetBoundingBox(X, Y, W, H, ...)` | 获取包围盒 |
| `SetRotationCenter(X, Y)` | 设置旋转中心 |
| `Rotate(Angle)` | 旋转 |
| `RotateEx(Angle, CX, CY)` | 绕指定点旋转 |
| `Skew(AngleX, AngleY)` | 倾斜 |
| `SkewEx(CX, CY, AngleX, AngleY)` | 绕指定点倾斜 |
| `Stretch(SX, [SY])` | 拉伸 |
| `StretchEx(CX, CY, SX, [SY])` | 绕指定点拉伸 |
| `Flip(Axes)` | 翻转 |
| `AffineTransform(a,b,c,d,e,f)` | 仿射变换 |
| `ClearTransformations()` | 清除变换 |
| `Project(...)` | 透视投影 |
| `Unproject()` | 移除透视投影 |

#### 对齐与分布

| 方法 | 说明 |
|------|------|
| `AlignToGrid()` | 对齐到网格 |
| `AlignToPage(Alignment)` | 对齐到页面 |
| `AlignToPageCenter(Alignment)` | 对齐到页面中心 |
| `AlignToPoint(Alignment, X, Y)` | 对齐到点 |
| `AlignToShape(Alignment, Shape)` | 对齐到形状 |
| `AlignToShapeRange(Alignment, ShapeRange)` | 对齐到形状范围 |
| `AlignRangeToGrid()` | 范围对齐到网格 |
| `AlignRangeToPage(Alignment)` | 范围对齐到页面 |
| `AlignRangeToPageCenter(Alignment)` | 范围对齐到页面中心 |
| `AlignRangeToPoint(Alignment, X, Y)` | 范围对齐到点 |
| `AlignRangeToShape(Alignment, Shape)` | 范围对齐到形状 |
| `AlignRangeToShapeRange(Alignment, SR)` | 范围对齐到另一范围 |
| `AlignAndDistribute(...)` | 对齐与分布 |
| `Distribute(Type, ...)` | 分布 |

#### 外观操作（批量）

| 方法 | 说明 |
|------|------|
| `ApplyFill(Fill)` | 应用已有的 Fill 对象 |
| `ApplyUniformFill(Color)` | 应用均匀色填充 |
| `ApplyFountainFill(StartColor, EndColor, ...)` | 应用渐变填充 |
| `ApplyHatchFill(...)` | 应用网格填充 |
| `ApplyCustomHatchFill(...)` | 应用自定义网格填充 |
| `ApplyPatternFill(Type, ...)` | 应用图案填充 |
| `ApplyPostscriptFill(...)` | 应用 PostScript 填充 |
| `ApplyTextureFill(Name, Library, ...)` | 应用纹理填充 |
| `ApplyNoFill()` | 移除填充 |
| `ApplyOutline(Outline)` | 应用已有的 Outline 对象 |
| `SetOutlineProperties(Width, Style, Color, ...)` | 批量设置轮廓属性 |
| `SetOutlinePropertiesEx(...)` | 扩展版轮廓属性设置 |
| `SetFillMode(Mode)` | 设置填充模式 |
| `SetPixelAlignedRendering(Value)` | 设置像素对齐渲染 |

#### 效果操作（批量）

| 方法 | 说明 |
|------|------|
| `ApplyEffectBCI(...)` | 亮度/对比度/强度 |
| `ApplyEffectColorBalance(...)` | 色彩平衡 |
| `ApplyEffectGamma(Gamma)` | 伽马 |
| `ApplyEffectHSL(...)` | 色相/饱和度/亮度 |
| `ApplyEffectInvert()` | 反色 |
| `ApplyEffectPosterize(Levels)` | 色调分离 |
| `ClearEffect(EffectType)` | 清除效果 |

#### 群组与合并

| 方法 | 说明 |
|------|------|
| `Group() As Shape` | 将范围内所有形状组合，返回组合 Shape |
| `Ungroup()` | 取消组合 |
| `UngroupAll()` | 递归取消所有嵌套组合 |
| `UngroupAllEx() As ShapeRange` | 递归取消并返回 ShapeRange |
| `UngroupEx() As ShapeRange` | 取消一层并返回 ShapeRange |
| `Combine() As Shape` | 合并为单一曲线形状 |
| `BreakApart()` | 拆分子路径 |
| `BreakApartEx() As ShapeRange` | 拆分并返回 ShapeRange |

#### 选择与层次

| 方法 | 说明 |
|------|------|
| `CreateSelection()` | 以本范围创建选区 |
| `AddToSelection()` | 将范围中的形状加入选区 |
| `RemoveFromSelection()` | 从选区中移除 |
| `Lock()` | 锁定全部形状 |
| `Unlock()` | 解锁全部形状 |
| `MoveToLayer(Layer)` | 移动到指定图层 |
| `CopyToLayer(Layer)` | 复制到指定图层 |
| `OrderToFront()` | 全部移到最前 |
| `OrderToBack()` | 全部移到最后 |
| `OrderForwardOne()` | 全部向前一层 |
| `OrderBackOne()` | 全部向后一层 |
| `OrderFrontOf(Shape)` | 全部移到指定形状前 |
| `OrderBackOf(Shape)` | 全部移到指定形状后 |
| `OrderReverse()` | 反转堆叠顺序 |

#### 其他

| 方法 | 说明 |
|------|------|
| `Delete()` | 删除范围中所有形状 |
| `Copy()` | 复制到剪贴板 |
| `Cut()` | 剪切到剪贴板 |
| `Duplicate([OffsetX], [OffsetY]) As ShapeRange` | 复制范围 |
| `Clone() As ShapeRange` | 克隆范围 |
| `StepAndRepeat(N, DX, DY, ...) As ShapeRange` | 步骤与重复 |
| `CopyPropertiesFrom(Source, Props)` | 从形状复制属性 |
| `ConvertToCurves()` | 全部转换为曲线 |
| `ConvertToBitmap(...) As Shape` | 全部转换为位图（旧版）|
| `ConvertToBitmapEx(...) As Shape` | 全部转换为位图（推荐）|
| `ConvertToSymbol(Name) As ShapeRange` | 转换为符号 |
| `ConvertOutlineToObject() As ShapeRange` | 将轮廓转换为对象 |
| `AddToPowerClip(Container, ...)` | 放入 PowerClip |
| `RemoveFromContainer([Level])` | 从 PowerClip 取出 |
| `CreateBoundary(X, Y, ...)` | 创建边界 |
| `CreateDocumentFrom() As Document` | 从范围创建新文档 |
| `CreateSelection()` | 创建选区 |
| `Fillet(Radius, ...)` | 圆角 |
| `Chamfer(A, B, ...)` | 倒角 |
| `Scallop(Radius, ...)` | 扇形角 |
| `EqualDivide(N, ...)` | 等分切割 |
| `GetLinkedShapes() As ShapeRange` | 获取关联形状 |
| `GetOverprintFillState() As cdrTriState` | 获取叠印填充状态 |
| `GetOverprintOutlineState() As cdrTriState` | 获取叠印轮廓状态 |
| `RestoreCloneLink(...)` | 恢复克隆链接 |
| `CustomCommand(Name, ...)` | 执行自定义命令 |

---

### 3.3 创建与管理 ShapeRange

```vba
' 方式 1：New 关键字手动创建并逐一添加形状
Sub ManualRange()
    Dim sr As New ShapeRange
    sr.Add ActiveLayer.CreateEllipse2(1, 1, 0.5)
    sr.Add ActiveLayer.CreateEllipse2(3, 1, 0.5)
    sr.Add ActiveLayer.CreateEllipse2(5, 1, 0.5)
    sr.ApplyUniformFill CreateRGBColor(255, 100, 0)
    Set sr(1).Fill.UniformColor = Nothing   ' 单独修改某一个
End Sub

' 方式 2：通过 FindShapes 获取
Sub FindRange()
    Dim sr As ShapeRange
    Set sr = ActivePage.FindShapes(Type:=cdrEllipseShape)
    MsgBox "页面中共有 " & sr.Count & " 个椭圆"
End Sub

' 方式 3：通过 Shapes.All 获取全部
Sub AllShapes()
    Dim sr As ShapeRange
    Set sr = ActivePage.Shapes.All
    sr.Lock   ' 锁定全部
End Sub

' 方式 4：通过选区获取
Sub FromSelection()
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    MsgBox "已选中 " & sr.Count & " 个形状"
End Sub

' 合并两个范围
Sub MergeRanges()
    Dim sr1 As ShapeRange, sr2 As ShapeRange
    Set sr1 = ActivePage.FindShapes(Type:=cdrRectangleShape)
    Set sr2 = ActivePage.FindShapes(Type:=cdrEllipseShape)
    sr1.AddRange sr2
    MsgBox "矩形和椭圆共 " & sr1.Count & " 个"
End Sub

' 从范围中移除非曲线形状（不删除，只移出范围）
Sub FilterToCurves()
    Dim sr As ShapeRange
    Dim i As Long
    Set sr = ActivePage.Shapes.All
    For i = sr.Count To 1 Step -1
        If sr(i).Type <> cdrCurveShape Then sr.Remove i
    Next i
    sr.CreateSelection
    MsgBox "已选中 " & sr.Count & " 个曲线形状"
End Sub
```

---

### 3.4 批量操作详解

```vba
' 批量对齐到页面中心（水平居中）
Sub CenterAllHorizontally()
    ActivePage.Shapes.All.AlignToPageCenter cdrAlignHCenter
End Sub

' 批量调整轮廓为 0.5pt 黑色
Sub SetOutlines()
    Dim sr As ShapeRange
    Set sr = ActivePage.Shapes.All
    sr.SetOutlineProperties 0.5 / 72 * 25.4, , CreateRGBColor(0, 0, 0)
End Sub

' 批量填充并组合
Sub FillAndGroup()
    Dim sr As New ShapeRange
    Dim n As Long
    For n = 1 To 8
        sr.Add ActiveLayer.CreateEllipse2(n * 1.2, 5, 0.5)
    Next n
    sr.ApplyFountainFill CreateRGBColor(255, 0, 0), _
                          CreateRGBColor(0, 0, 255), _
                          cdrLinearFountainFill
    Dim grp As Shape
    Set grp = sr.Group
    grp.SetPositionEx cdrCenter, ActivePage.SizeWidth / 2, ActivePage.SizeHeight / 2
End Sub

' 按 CQL 排序（从上到下、从左到右）
Sub SortTopToBottom()
    Dim sr As ShapeRange
    Set sr = ActivePage.Shapes.All
    sr.Sort "@shape1.Top * 1000 - @shape1.Left > @shape2.Top * 1000 - @shape2.Left"
    ' 重新排列形状
    Dim s As Shape
    Dim y As Double
    y = ActivePage.SizeHeight - 0.5
    For Each s In sr
        s.SetPositionEx cdrCenter, 1, y
        y = y - 0.5
    Next s
End Sub

' 等分切割
Sub DivideShapes()
    Dim s As Shape
    Set s = ActiveLayer.CreateRectangle2(1, 1, 6, 4)
    s.Fill.UniformColor.RGBAssign 100, 200, 100
    Dim pieces As ShapeRange
    Set pieces = s.EqualDivide(3)   ' 水平切割为 3 份
    pieces.ApplyFountainFill CreateRGBColor(255, 0, 0), _
                              CreateRGBColor(0, 255, 0), _
                              cdrRadialFountainFill
End Sub
```

---

## 四、Shapes 集合

**描述**：`Shapes` 是一个固定集合，代表某个容器（页面、图层、组合形状）中的所有形状。与 `ShapeRange` 的区别是：`Shapes` 反映的是文档的真实结构，不可直接增删成员；而 `ShapeRange` 是临时的工作集合，可手动增删。

### 4.1 Shapes 属性与方法一览

| 成员 | 类型 | 说明 |
|------|------|------|
| `Count` | `Long`（只读）| 形状数量 |
| `Item(IndexOrName)` | `Shape`（只读）| 按索引或名称取形状（默认属性）|
| `First` | `Shape`（只读）| 第一个形状（Z 轴最低）|
| `Last` | `Shape`（只读）| 最后一个形状（Z 轴最高）|
| `Parent` | `Object`（只读）| 父对象（页面/图层/组合）|
| `Application` | `Application`（只读）| 顶层应用 |
| `All()` | `ShapeRange` | 返回包含全部形状的 ShapeRange |
| `AllExcluding(IndexArray)` | `ShapeRange` | 返回排除指定索引后的 ShapeRange |
| `Range(IndexArray)` | `ShapeRange` | 返回指定索引子集的 ShapeRange |
| `FindShape([Name], [Type], [StaticID], [Recursive], [Query])` | `Shape` | 查找单个形状 |
| `FindShapes([Name], [Type], [Recursive], [Query])` | `ShapeRange` | 查找多个形状 |

### 4.2 查找与筛选详解

```vba
' 按名称查找
Sub FindByName()
    Dim s As Shape
    Set s = ActivePage.FindShape(Name:="Logo")
    If s Is Nothing Then
        MsgBox "未找到名为 Logo 的形状"
    Else
        s.CreateSelection
    End If
End Sub

' 按类型查找所有椭圆
Sub FindAllEllipses()
    Dim sr As ShapeRange
    Set sr = ActivePage.FindShapes(Type:=cdrEllipseShape)
    sr.ApplyUniformFill CreateRGBColor(255, 200, 0)
End Sub

' 按 CQL 查询查找（宽度大于 3 英寸的矩形）
Sub FindByQuery()
    Dim sr As ShapeRange
    Set sr = ActivePage.FindShapes( _
        Type:=cdrRectangleShape, _
        Query:="@this.width > 3")
    MsgBox "找到 " & sr.Count & " 个宽度超过3英寸的矩形"
End Sub

' 按静态 ID 查找（用于脚本中跨操作稳定引用）
Sub FindByStaticID()
    Dim targetID As Long
    ' 记录 ID
    targetID = ActiveShape.StaticID
    ' 稍后查找
    Dim s As Shape
    Set s = ActivePage.FindShape(StaticID:=targetID)
End Sub

' 获取图层内第 2 到第 5 个形状
Sub GetSubset()
    Dim sr As ShapeRange
    Set sr = ActiveLayer.Shapes.Range(Array(2, 3, 4, 5))
    sr.CreateSelection
End Sub

' 获取除最后一个形状外的所有形状
Sub AllButLast()
    Dim ss As Shapes
    Set ss = ActivePage.Shapes
    Dim sr As ShapeRange
    Set sr = ss.AllExcluding(Array(ss.Count))
    sr.Delete
End Sub

' 用 For Each 遍历 Shapes 集合
Sub IterateShapes()
    Dim s As Shape
    For Each s In ActivePage.Shapes
        Debug.Print s.Name & "  类型：" & s.Type & _
            "  位置：(" & s.CenterX & ", " & s.CenterY & ")"
    Next s
End Sub
```

---

## 五、SelectionInfo 对象

**描述**：`SelectionInfo` 对象提供当前选区的**状态信息**，通过 `Document.SelectionInfo` 访问。它的属性全部是只读的，主要用于判断当前选区包含哪类对象、是否可以进行某种操作等。

### 5.1 SelectionInfo 属性一览

#### 基础信息

| 属性 | 类型 | 说明 |
|------|------|------|
| `Count` | `Long` | 当前选中对象的数量 |
| `FirstShape` | `Shape` | 选区中第一个形状 |
| `SecondShape` | `Shape` | 选区中第二个形状 |
| `Parent` | `Document` | 父文档 |

#### "能否操作" 判断属性（`Can...`）

| 属性 | 说明 |
|------|------|
| `CanApplyBlend` | 是否可以应用调和 |
| `CanApplyContour` | 是否可以应用轮廓图 |
| `CanApplyDistortion` | 是否可以应用变形 |
| `CanApplyEnvelope` | 是否可以应用封套 |
| `CanApplyFill` | 是否可以应用填充 |
| `CanApplyFillOutline` | 是否可以应用填充和轮廓 |
| `CanApplyOutline` | 是否可以应用轮廓 |
| `CanApplyTransparency` | 是否可以应用透明度 |
| `CanAssignURL` | 是否可以分配 URL |
| `CanClone` | 是否可以克隆 |
| `CanCloneBlend` | 是否可以克隆调和 |
| `CanCloneContour` | 是否可以克隆轮廓图 |
| `CanCloneDropShadow` | 是否可以克隆阴影 |
| `CanCloneExtrude` | 是否可以克隆立体化 |
| `CanCopyBlend` | 是否可以复制调和 |
| `CanCopyContour` | 是否可以复制轮廓图 |
| `CanCopyDistortion` | 是否可以复制变形 |
| `CanCopyDropShadow` | 是否可以复制阴影 |
| `CanCopyEnvelope` | 是否可以复制封套 |
| `CanCopyExtrude` | 是否可以复制立体化 |
| `CanCopyLens` | 是否可以复制透镜 |
| `CanCopyPerspective` | 是否可以复制透视 |
| `CanCopyPowerclip` | 是否可以复制 PowerClip |
| `CanCreateBlend` | 是否可以创建调和（需选中两个可混合形状）|
| `CanDeleteControl` | 是否可以删除控制形状 |
| `CanLockShapes` | 是否可以锁定形状 |
| `CanPrint` | 是否可以打印 |
| `CanUngroup` | 是否可以取消组合 |
| `CanUnlockShapes` | 是否可以解锁 |

#### "选区包含什么" 判断属性（`Is...`）

| 属性 | 说明 |
|------|------|
| `IsArtisticTextSelected` | 是否选中了美工字 |
| `IsParagraphTextSelected` | 是否选中了段落文字 |
| `IsTextSelected` | 是否选中了任何文字 |
| `IsTextSelection` | 选区是否是文字内部的文本选区 |
| `IsBitmapPresent` | 是否有位图 |
| `IsBitmapSelected` | 是否选中了位图 |
| `IsExternalBitmapSelected` | 是否选中了外链位图 |
| `IsNonExternalBitmapSelected` | 是否选中了嵌入位图 |
| `IsMaskedBitmapPresent` | 是否有蒙版位图 |
| `IsGroupSelected` | 是否选中了组合 |
| `IsGroup` | 是否是单一组合 |
| `IsBlendGroup` | 是否是调和组 |
| `IsBlendControl` | 是否是调和控制形状 |
| `IsContourGroup` | 是否是轮廓图组 |
| `IsContourControl` | 是否是轮廓图控制形状 |
| `IsDropShadowGroup` | 是否是阴影组 |
| `IsDropShadowControl` | 是否是阴影控制形状 |
| `IsExtrudeGroup` | 是否是立体组 |
| `IsExtrudeControl` | 是否是立体控制形状 |
| `IsDistortion` | 是否是变形 |
| `IsDistortionPresent` | 选区中是否存在变形效果 |
| `IsEnvelope` | 是否是封套 |
| `IsEnvelopePresent` | 选区中是否存在封套 |
| `IsPerspective` | 是否是透视 |
| `IsPerspectivePresent` | 选区中是否存在透视 |
| `IsMeshFillPresent` | 是否存在网格填充 |
| `IsMeshFillSelected` | 是否选中了网格填充 |
| `IsOLESelected` | 是否选中了 OLE 对象 |
| `IsConnector` | 是否是连接器 |
| `IsConnectorLine` | 是否是连接线 |
| `IsConnectorLineSelected` | 是否选中了连接线 |
| `IsConnectorSelected` | 是否选中了连接器 |
| `IsRegularShape` | 是否是规则形状（矩形/椭圆/多边形）|
| `IsGuidelineSelected` | 是否选中了辅助线 |
| `IsFittedText` | 是否是路径文字 |
| `IsFittedTextSelected` | 是否选中了路径文字 |
| `IsFittedTextControl` | 是否是路径文字的控制路径 |
| `IsRollOverSelected` | 是否选中了翻转对象 |
| `IsEditingRollOver` | 是否正在编辑翻转 |
| `IsEditingText` | 是否正在编辑文字 |
| `IsBevelGroup` | 是否是斜面组 |
| `IsCloneControl` | 是否是克隆控制形状 |
| `IsControlSelected` | 是否选中了控制形状 |
| `IsControlShape` | 是否是控制形状 |
| `IsNaturalMediaControl` | 是否是自然媒体控制形状 |
| `IsNaturalMediaGroup` | 是否是自然媒体组 |
| `IsOnPowerClipContents` | 是否在 PowerClip 内容中 |
| `IsAttachedToDimension` | 是否附加到标注 |
| `IsDimensionControl` | 是否是标注控制形状 |
| `IsLinkControlSelected` | 是否选中了链接控制 |
| `IsLinkGroupSelected` | 是否选中了链接组 |
| `IsSoundObjectSelected` | 是否选中了声音对象 |
| `IsInternetObjectSelected` | 是否选中了 Internet 对象 |
| `IsSecondContourControl` | 是否是第二个轮廓图控制 |
| `IsSecondDropShadowControl` | 是否是第二个阴影控制 |
| `IsSecondExtrudeControl` | 是否是第二个立体化控制 |
| `IsSecondNaturalMediaControl` | 是否是第二个自然媒体控制 |

#### 关联形状访问属性

| 属性 | 说明 |
|------|------|
| `BlendBottomShape` | 调和的底部形状 |
| `BlendTopShape` | 调和的顶部形状 |
| `BlendPath` | 调和路径形状 |
| `ContourControlShape` | 轮廓图控制形状 |
| `ContourGroup` | 轮廓图组 |
| `DistortionShape` | 变形形状 |
| `DistortionType` | 变形类型 |
| `DropShadowControlShape` | 阴影控制形状 |
| `DropShadowGroup` | 阴影组 |
| `ExtrudeGroup` | 立体组 |
| `ExtrudeFaceShape` | 立体面形状 |
| `ExtrudeBevelGroup` | 立体斜面组 |
| `DimensionControlShape` | 标注控制形状 |
| `DimensionGroup` | 标注组 |
| `FittedText` | 路径文字形状 |
| `FittedTextControlShape` | 路径文字的控制路径 |
| `NaturalMediaControlShape` | 自然媒体控制形状 |
| `NaturalMediaGroup` | 自然媒体组 |
| `ConnectorLines` | 连接线集合 |
| `ContainsRollOverParent` | 是否包含翻转父对象 |
| `FirstShapeWithFill` | 第一个有填充的形状 |
| `FirstShapeWithOutline` | 第一个有轮廓的形状 |
| `HasAutoLabelText` | 是否有自动标签文字 |

---

### 5.2 实用示例

```vba
' 显示选区摘要信息
Sub ShowSelectionSummary()
    Dim info As SelectionInfo
    Set info = ActiveDocument.SelectionInfo
    If info.Count = 0 Then
        MsgBox "当前无选中对象"
        Exit Sub
    End If
    Dim s As String
    s = "选中对象数：" & info.Count & vbCr
    s = s & "包含位图：" & info.IsBitmapPresent & vbCr
    s = s & "包含文字：" & info.IsTextSelected & vbCr
    s = s & "包含组合：" & info.IsGroupSelected & vbCr
    s = s & "可以填充：" & info.CanApplyFill & vbCr
    s = s & "可以调和：" & info.CanCreateBlend & vbCr
    MsgBox s
End Sub

' 根据选区类型执行不同操作
Sub SmartOperation()
    Dim info As SelectionInfo
    Set info = ActiveDocument.SelectionInfo
    If info.Count = 0 Then Exit Sub

    If info.CanApplyFill Then
        ' 对可填充对象应用红色填充
        ActiveSelectionRange.ApplyUniformFill CreateRGBColor(255, 0, 0)
    End If

    If info.CanCreateBlend And info.Count = 2 Then
        ' 对两个形状创建调和
        info.FirstShape.CreateBlend info.SecondShape
    End If

    If info.IsGroupSelected Then
        ' 取消选中的组合
        ActiveSelectionRange.Ungroup
    End If
End Sub
```

---

## 六、选择操作（Selection API）

### 选区相关 API 汇总

| API | 所在对象 | 说明 |
|-----|---------|------|
| `ActiveShape` | `Application` | 单个激活形状（只读）|
| `ActiveSelection` | `Application` | 选区 Shape（类型为 `cdrSelectionShape`，只读）|
| `ActiveSelectionRange` | `Application` | 选区 ShapeRange（只读）|
| `Document.Selection()` | `Document` | 返回选区 Shape（同 `ActiveSelection`）|
| `Document.SelectionRange` | `Document` | 返回选区 ShapeRange（只读）|
| `Document.SelectionInfo` | `Document` | 返回 `SelectionInfo` 对象（只读）|
| `Document.AddToSelection(...)` | `Document` | 将形状数组加入选区 |
| `Document.CreateSelection(...)` | `Document` | 从形状数组创建新选区 |
| `Document.ClearSelection()` | `Document` | 清除当前选区 |
| `Document.RemoveFromSelection(...)` | `Document` | 从选区中移除指定形状 |
| `Shape.Selected` | `Shape` | 读写：是否选中 |
| `Shape.AddToSelection()` | `Shape` | 加入选区 |
| `Shape.RemoveFromSelection()` | `Shape` | 移出选区 |
| `Shape.CreateSelection()` | `Shape` | 创建仅含本形状的选区 |
| `ShapeRange.AddToSelection()` | `ShapeRange` | 将范围加入选区 |
| `ShapeRange.RemoveFromSelection()` | `ShapeRange` | 从选区中移除范围 |
| `ShapeRange.CreateSelection()` | `ShapeRange` | 以范围创建选区 |

### 常用选择操作示例

```vba
' 全选当前页面所有形状
Sub SelectAll()
    ActivePage.Shapes.All.CreateSelection
End Sub

' 清除选区
Sub Deselect()
    ActiveDocument.ClearSelection
End Sub

' 反选：选中当前未被选中的形状
Sub InvertSelection()
    Dim s As Shape
    For Each s In ActivePage.Shapes
        s.Selected = Not s.Selected
    Next s
End Sub

' 只选中所有矩形
Sub SelectRects()
    ActiveDocument.ClearSelection
    Dim sr As ShapeRange
    Set sr = ActivePage.FindShapes(Type:=cdrRectangleShape)
    sr.CreateSelection
End Sub

' 从选区中移除文字
Sub DeselectText()
    ActivePage.FindShapes(Type:=cdrTextShape).RemoveFromSelection
End Sub

' 获取选区中心并在该位置创建标记点
Sub MarkSelectionCenter()
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then Exit Sub
    Dim cx As Double, cy As Double
    cx = sr.CenterX
    cy = sr.CenterY
    Dim marker As Shape
    Set marker = ActiveLayer.CreateEllipse2(cx, cy, 0.1)
    marker.Fill.UniformColor.RGBAssign 255, 0, 0
    marker.Outline.Type = cdrNoOutline
End Sub

' 用 ActiveSelection.Shapes 遍历所有选中形状
Sub ProcessSelected()
    Dim s As Shape
    For Each s In ActiveSelection.Shapes
        Debug.Print "选中：" & s.Name & " 类型：" & s.Type
    Next s
End Sub

' 保存和恢复选区
Sub SaveRestoreSelection()
    ' 保存
    Dim saved As New ShapeRange
    Dim s As Shape
    For Each s In ActiveSelectionRange
        saved.Add s
    Next s
    ' 做一些其他操作...
    ActiveDocument.ClearSelection
    ' 恢复
    saved.CreateSelection
End Sub
```

---

## 七、cdrShapeType 枚举常量

| 常量 | 值 | 说明 |
|------|----|------|
| `cdrNoShape` | 0 | 无形状 |
| `cdrRectangleShape` | 1 | 矩形 |
| `cdrEllipseShape` | 2 | 椭圆 |
| `cdrCurveShape` | 3 | 曲线 |
| `cdrPolygonShape` | 4 | 多边形（星形）|
| `cdrBitmapShape` | 5 | 位图 |
| `cdrTextShape` | 6 | 文字 |
| `cdrGroupShape` | 7 | 组合 |
| `cdrSelectionShape` | 8 | 选区 |
| `cdrGuidelineShape` | 9 | 辅助线 |
| `cdrBlendGroupShape` | 10 | 调和组 |
| `cdrExtrudeGroupShape` | 11 | 立体化组 |
| `cdrOLEObjectShape` | 12 | OLE 对象 |
| `cdrContourGroupShape` | 13 | 轮廓图组 |
| `cdrLinearDimensionShape` | 14 | 线性标注 |
| `cdrBevelGroupShape` | 15 | 斜面组 |
| `cdrDropShadowGroupShape` | 16 | 阴影组 |
| `cdr3DObjectShape` | 17 | 3D 对象 |
| `cdrArtisticMediaGroupShape` | 18 | 自然媒体组 |
| `cdrConnectorShape` | 19 | 连接器 |
| `cdrMeshFillShape` | 20 | 网格填充 |
| `cdrCustomShape` | 21 | 自定义形状（表格、QR 码等）|
| `cdrCustomEffectGroupShape` | 22 | 自定义效果组 |
| `cdrSymbolShape` | 23 | 符号 |
| `cdrHTMLFormObjectShape` | 24 | HTML 表单对象 |
| `cdrHTMLActiveObjectShape` | 25 | HTML 活动对象 |

---

## 八、综合实战示例

### 示例 1：将所有形状分类并统计报告

```vba
Sub ShapeReport()
    Dim counts(25) As Long
    Dim s As Shape
    Dim total As Long
    Dim report As String
    Dim typeNames(25) As String
    typeNames(1) = "矩形": typeNames(2) = "椭圆"
    typeNames(3) = "曲线": typeNames(4) = "多边形"
    typeNames(5) = "位图": typeNames(6) = "文字"
    typeNames(7) = "组合": typeNames(9) = "辅助线"
    typeNames(21) = "自定义(含表格)"

    For Each s In ActivePage.Shapes
        If s.Type >= 0 And s.Type <= 25 Then
            counts(s.Type) = counts(s.Type) + 1
            total = total + 1
        End If
    Next s
    report = "当前页面形状统计（共 " & total & " 个）：" & vbCr & vbCr
    Dim i As Integer
    For i = 1 To 25
        If counts(i) > 0 Then
            Dim n As String
            If typeNames(i) <> "" Then n = typeNames(i) Else n = "类型" & i
            report = report & n & "：" & counts(i) & " 个" & vbCr
        End If
    Next i
    MsgBox report
End Sub
```

### 示例 2：将选区中的形状排成均匀网格

```vba
Sub ArrangeInGrid()
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then MsgBox "请先选中形状": Exit Sub

    Dim cols As Long
    cols = CLng(InputBox("每行显示几列？", "网格排列", "4"))
    If cols <= 0 Then Exit Sub

    Dim cellW As Double, cellH As Double
    cellW = 2: cellH = 2    ' 单元格尺寸（文档单位）
    Dim i As Long, row As Long, col As Long
    For i = 1 To sr.Count
        row = (i - 1) \ cols
        col = (i - 1) Mod cols
        sr(i).SetPositionEx cdrCenter, col * cellW + cellW / 2, -(row * cellH) + ActivePage.SizeHeight - cellH / 2
    Next i
End Sub
```

### 示例 3：批量替换填充颜色（用 CQL 查询）

```vba
Sub ReplaceFillColor()
    ' 查找所有红色填充的矩形
    Dim sr As ShapeRange
    Set sr = ActivePage.FindShapes( _
        Type:=cdrRectangleShape, _
        Query:="@this.fill.type = 'uniform' and @this.fill.color.r > 200")
    If sr.Count = 0 Then
        MsgBox "未找到红色矩形"
        Exit Sub
    End If
    sr.ApplyUniformFill CreateRGBColor(0, 120, 200)    ' 替换为蓝色
    MsgBox "已将 " & sr.Count & " 个红色矩形改为蓝色"
End Sub
```

### 示例 4：逐层提取所有嵌套形状并展平

```vba
Sub FlattenAll()
    ' 将所有组合取消，使所有形状都在当前图层
    Dim changed As Boolean
    Do
        changed = False
        Dim s As Shape
        For Each s In ActivePage.Shapes
            If s.Type = cdrGroupShape Then
                s.UngroupAll
                changed = True
                Exit For
            End If
        Next s
    Loop While changed
    MsgBox "所有嵌套组合已展平，当前页面形状数：" & ActivePage.Shapes.Count
End Sub
```

### 示例 5：用 ShapeRange 构建注册标记

```vba
Sub CreateRegistrationMark()
    Dim sr As New ShapeRange
    ' 外圆
    sr.Add ActiveLayer.CreateEllipse2(0, 0, 0.5)
    ' 内圆
    sr.Add ActiveLayer.CreateEllipse2(0, 0, 0.15)
    ' 水平线
    sr.Add ActiveLayer.CreateLineSegment(-0.6, 0, 0.6, 0)
    ' 垂直线
    sr.Add ActiveLayer.CreateLineSegment(0, -0.6, 0, 0.6)
    ' 设置轮廓
    sr.SetOutlineProperties 0.25 / 72 * 25.4, , CreateRGBColor(0, 0, 0)
    sr(2).Outline.Type = cdrNoOutline   ' 内圆无轮廓
    sr(2).Fill.UniformColor.RGBAssign 0, 0, 0  ' 内圆黑色填充
    ' 组合
    Dim mark As Shape
    Set mark = sr.Group
    mark.Name = "套准标记"
    ' 移到页面右下角
    mark.SetPositionEx cdrCenter, ActivePage.SizeWidth - 1, 1
End Sub
```

---

## 附：对象与方法快速对照

| 需求 | 推荐 API |
|------|---------|
| 访问单个选中形状 | `ActiveShape` |
| 访问整个选区（可操作）| `ActiveSelection` 或 `ActiveDocument.Selection()` |
| 获取选区中的全部形状 | `ActiveSelection.Shapes` |
| 获取选区 ShapeRange | `ActiveSelectionRange` 或 `ActiveDocument.SelectionRange` |
| 查询选区状态 | `ActiveDocument.SelectionInfo` |
| 找所有某类型形状 | `ActivePage.FindShapes(Type:=cdrXxxShape)` |
| 找所有形状（转 ShapeRange）| `ActivePage.Shapes.All` |
| 手动构建形状集合 | `Dim sr As New ShapeRange` + `sr.Add ...` |
| 批量操作多个形状 | `ShapeRange` 的各种方法 |
| 遍历容器中的所有形状 | `For Each s In Page/Layer/Group.Shapes` |
| 遍历 ShapeRange | `For Each s In sr` 或 `For i = 1 To sr.Count` |
| 按名称访问形状 | `ActivePage.FindShape(Name:="xxx")` |
| 按索引访问形状 | `ActivePage.Shapes(n)` 或 `sr(n)` |
