# CorelDRAW 脚本创建表格示例

本目录包含一个使用 **CorelDRAW X7 COM API / VBA 宏** 在 CorelDRAW 中动态创建表格的完整示例脚本。

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `CreateTable示例.vba` | VBA 宏脚本主体，可直接复制到 CorelDRAW 宏编辑器中运行 |
| `CorelDRAW X7 脚本手册.chm` | CorelDRAW X7 官方脚本参考手册（CHM 格式） |
| `CorelDraw脚本参考手册X7_0*.pdf` | CorelDRAW X7 脚本参考手册分册（PDF 格式，共 8 册） |

---

## 运行环境

- **软件**：CorelDRAW X7 或更高版本（X8、2017、2018、2019、2020、2021）
- **语言**：VBA（Visual Basic for Applications）
- **权限**：无需安装额外插件，使用 CorelDRAW 内置宏引擎即可

---

## 快速开始

1. 打开 CorelDRAW。
2. 点击菜单 **工具（Tools）→ 宏（Macros）→ 宏编辑器（Macro Editor）**，  
   或直接按快捷键 **Alt+F11**。
3. 在左侧"工程"窗格中，选中 `GlobalMacros` 或任意模块，右键 **插入 > 模块**。
4. 将 `CreateTable示例.vba` 的全部内容粘贴到新模块中。
5. 将光标置于 `Sub Main()` 内，按 **F5** 运行，或点击工具栏的运行按钮。
6. 切换回 CorelDRAW 主窗口，即可看到生成的表格。

---

## 脚本功能概览

```
Main()
├── 若无活动文档则自动新建
├── 在当前图层调用 CreateTable() 创建 5×4 表格（120 mm × 80 mm）
├── FormatTableHeader()  ── 设置表头样式（深蓝背景、白色加粗文字）
├── FillTableContent()   ── 填充列标题与 4 行产品数据，奇偶行交替底色
└── SetColumnWidths()    ── 按比例（1:2:1.5:1.5）分配各列宽度
```

生成的表格效果示意：

| 产品编号 | 产品名称 | 单价（元） | 库存（件） |
|---------|---------|-----------|-----------|
| P-001   | 圆珠笔   | 2.50      | 1200      |
| P-002   | 笔记本   | 15.00     | 560       |
| P-003   | 文件夹   | 8.80      | 340       |
| P-004   | 订书机   | 32.00     | 88        |

---

## 关键 API 说明

### 1. 创建表格

```vba
' layer.CreateTable(行数, 列数, 宽度_磅, 高度_磅) → Shape
Dim tblShape As Shape
Set tblShape = layer.CreateTable(5, 4, _
                    MillimetersToPoints(120), _
                    MillimetersToPoints(80))
```

> CorelDRAW 内部坐标单位为**磅（point）**，1 mm ≈ 2.8346 pt。

### 2. 定位表格

```vba
' 将表格左上角移动到页面指定位置
tblShape.SetPosition MillimetersToPoints(30), MillimetersToPoints(170)
```

### 3. 访问单元格

```vba
Dim tbl  As Table
Dim cell As Cell
Set tbl  = tblShape.Table
Set cell = tbl.Cell(1, 1)     ' Cell(行, 列)，下标从 1 开始
```

### 4. 填写文字

```vba
cell.Text.Story.TextRange.Text = "Hello"
```

### 5. 设置背景色

```vba
cell.Fill.UniformColor.RGBAssign 31, 73, 125   ' R, G, B
```

### 6. 设置字体样式

```vba
With cell.Text.Story.TextRange.Font
    .Bold = True
    .Size = 10          ' 磅
End With
```

### 7. 文字对齐方式

```vba
' 常用常量：cdrCenterAlignment / cdrLeftAlignment / cdrRightAlignment
cell.Text.Story.TextRange.Alignment = cdrCenterAlignment
```

### 8. 设置列宽

```vba
tbl.Column(1).Width = MillimetersToPoints(30)
```

---

## 扩展思路

- **合并单元格**：`tbl.Cell(1,1).Merge tbl.Cell(1,4)` 可将第 1 行全部列合并为标题行。
- **设置边框**：通过 `cell.Outline` 对象控制边框颜色和线宽。
- **从 Excel 读取数据**：结合 `CreateObject("Excel.Application")` 读取工作表数据后填入表格。
- **循环批量生成**：在循环中重复调用 `CreateTable`，可批量生成多张样式一致的表格。

---

## 参考资料

- 本仓库中的《CorelDraw脚本参考手册X7》（PDF 1~8 册）
- 本仓库中的《CorelDRAW X7 脚本手册》（CHM）
- CorelDRAW 官方开发者文档：https://community.coreldraw.com/sdk/
