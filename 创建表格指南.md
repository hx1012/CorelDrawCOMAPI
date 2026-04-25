# CorelDRAW 脚本创建表格示例

本目录包含通过阅读官方《CorelDRAW X7 脚本参考手册》（PDF/CHM）整理的 VBA 宏示例，演示如何用 **CorelDRAW X7 COM API** 动态创建和格式化表格。

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `CreateTable示例.vba` | VBA 宏脚本，含两个可运行示例（产品表 + 日历） |
| `CorelDRAW X7 脚本手册.chm` | CorelDRAW X7 官方脚本参考手册（CHM 格式） |
| `CorelDraw脚本参考手册X7_0*.pdf` | CorelDRAW X7 脚本参考手册分册（PDF 格式，共 8 册） |

---

## 运行环境

- **软件**：CorelDRAW X7 或更高版本（X8、2017–2021）
- **语言**：VBA（Visual Basic for Applications）
- **权限**：无需安装额外插件，使用 CorelDRAW 内置宏引擎即可

---

## 快速开始

1. 打开 CorelDRAW。
2. 点击菜单 **工具（Tools）→ 宏（Macros）→ 宏编辑器（Macro Editor）**，或按 **Alt+F11**。
3. 在左侧"工程"窗格中插入新模块，将 `CreateTable示例.vba` 全部内容粘贴进去。
4. 将光标置于 `Sub Main()` 或 `Sub CreateCalendar_January2009()` 内，按 **F5** 运行。
5. 切换回 CorelDRAW 主窗口查看结果。

---

## 脚本功能概览

脚本包含两个示例：

### 示例一：`Main` — 产品信息表

```
CreateCustomShape("Table", 20, 20, 140, 70, 4, 5)
├── 第 1 行：列标题（深蓝背景、10 号字、居中）
├── 第 2-5 行：产品数据（斑马纹、数字列右对齐）
├── 列宽比例 1:2:1.5:1.5（通过 SetWidth 设置）
└── 整体外框线宽 0.5 mm
```

生成效果：

| 产品编号 | 产品名称 | 单价（元） | 库存（件） |
|---------|---------|-----------|-----------|
| P-001   | 圆珠笔   |      2.50 |      1200 |
| P-002   | 笔记本   |     15.00 |       560 |
| P-003   | 文件夹   |      8.80 |       340 |
| P-004   | 订书机   |     32.00 |        88 |

### 示例二：`CreateCalendar_January2009` — 2009 年 1 月日历

官方手册 `Layer.CreateCustomShape` 章节中的经典示例（原文忠实复现）：

```
CreateCustomShape("Table", 1, 10, 5, 7, 7, 6)
├── 填入星期缩写（第 1 行）
├── AddRow 1 — 在顶部插入月份标题行
├── 合并标题行全部 7 格，写入 "January"，字号 22，居中
├── 按顺序填入日期 1~31（从单元格编号 13 开始）
├── 合并空白格（编号 9~12）并填灰色
└── 为 1 月 1 日加绿色高亮边框
```

---

## 核心 API 速查（来自官方文档）

### 1. 创建表格

```vba
' 语法（官方文档原文）：
' CreateCustomShape("Table", Left, Top, Right, Bottom, Columns, Rows) As Shape
' 参数均为文档单位（默认毫米）；Left/Top/Right/Bottom 是表格四边到页面边框的距离

Set s = ActiveLayer.CreateCustomShape("Table", 20, 20, 140, 70, 4, 5)
```

> ⚠️ 不存在 `CreateTable()` 方法。必须通过 `CreateCustomShape("Table", ...)` 创建。

### 2. 访问 TableShape 对象

```vba
Dim ts As Object
Set ts = s.Custom   ' Shape.Custom 返回 TableShape
```

### 3. 访问单元格

```vba
' Cell(Column, Row) —— 列索引在前，行索引在后，均从 1 开始
Dim cell As Object
Set cell = ts.Cell(2, 3)   ' 第 2 列、第 3 行的单元格
```

### 4. 写入文字

```vba
cell.TextShape.Text.Story = "Hello"
' 设置字号
cell.TextShape.Text.Story.Words.All.Size = 12
' 设置对齐
cell.TextShape.Text.Story.Alignment = cdrCenterAlignment
```

### 5. 填充背景色

```vba
Dim f As Fill
Set f = ActiveDocument.CreateFill("MyFill")
f.ApplyUniformFill CreateRGBColor(31, 73, 125)
cell.Fill.ApplyUniformFill CreateRGBColor(31, 73, 125)

' 对一个范围批量填充（CellRange(ColStart, RowStart, ColEnd, RowEnd)）
ts.CellRange(1, 1, 4, 1).ApplyFill f
```

### 6. 合并单元格

```vba
' 方式一：通过行的所有单元格
ts.Rows(1).Cells.All.Merge

' 方式二：通过 CellRange
ts.CellRange(1, 1, 4, 1).Merge

' 方式三：通过 Cells 索引范围
ts.Cells.Range(9, 10, 11, 12).Merge
```

### 7. 设置边框

```vba
' 整张表外框（通过 Shape.Outline）
s.Outline.Width = 0.5

' 某个范围的边框
ts.Rows(1).Cells.All.Borders.All.Width = 0.05

' 特定单元格边框颜色（绿色）
ts.Cells.Range(10).Borders.All.Color.RGBAssign 0, 255, 0
```

### 8. 设置行高 / 列宽

```vba
' SetWidth(Width, ResizeTable)
ts.Columns(1).SetWidth 20, False   ' 设置第 1 列宽 20 mm，不整体缩放

' SetHeight(Height, ResizeTable)
ts.Rows(2).SetHeight 10, False     ' 设置第 2 行高 10 mm

' 也可直接赋值（Width 属性）
ts.Columns(1).Width = 20
```

### 9. 添加 / 删除行列

```vba
ts.AddRow 1         ' 在第 1 行前插入新行
ts.AddColumn        ' 在最后追加一列
ts.Rows(3).Delete   ' 删除第 3 行
```

### 10. 按顺序索引单元格（Cells 集合）

```vba
' Cells 从左到右、从上到下连续编号（从 1 开始）
' 已合并的一组单元格只占 1 个编号
ts.Cells(13).TextShape.Text.Story = 1   ' 第 13 个单元格写入 "1"
```

---

## 重要注意事项

| 错误用法 ❌ | 正确用法 ✅ | 来源 |
|-----------|-----------|------|
| `layer.CreateTable(rows, cols, w, h)` | `layer.CreateCustomShape("Table", L, T, R, B, cols, rows)` | 手册 Layer.CreateCustomShape 章节 |
| `shape.Table` | `shape.Custom` | 手册 TableShape 类说明 |
| `cell.Text.Story.TextRange.Text` | `cell.TextShape.Text.Story` | 手册 TableCell.TextShape 属性 |
| `tbl.Column(i).Width = v` | `ts.Columns(i).SetWidth v, False` 或 `.Width = v` | 手册 TableColumn.SetWidth 方法 |
| `Cell(Row, Col)` | `Cell(Column, Row)` — 列在前！ | 手册 TableShape.Cell 方法 |

---

## 参考资料

- 本仓库中的《CorelDraw脚本参考手册X7》PDF（共 8 册）
  - 表格创建：第 3 册 第 559-560 页（`Layer.CreateCustomShape`）
  - TableShape 类：第 5 册 第 601-617 页
  - TableCell 类：第 5 册 第 497-531 页
  - TableBorders 类：第 5 册 第 484-496 页
  - cdrTableBorder 枚举：第 1 册
- 本仓库中的《CorelDRAW X7 脚本手册》（CHM）
- CorelDRAW 官方开发者社区：https://community.coreldraw.com/sdk/

