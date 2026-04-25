# CorelDRAW X7 对象数据（Object Data）API 完整指南

> 内容来源：《CorelDRAW X7 脚本参考手册》第 1–6 册（本仓库 PDF）  
> 涵盖对象：**DataFields · DataField · DataItems · DataItem**  
> 入口属性：**Shape.ObjectData · Shape.Properties · Document.DataFields**

---

## 目录

1. [概念背景与应用场景](#一概念背景与应用场景)
2. [对象模型结构](#二对象模型结构)
3. [Shape.Properties 与 Shape.ObjectData 的区别](#三shapeproperties-与-shapeobjectdata-的区别)
4. [Document.DataFields — 文档级字段定义](#四documentdatafields--文档级字段定义)
   - 4.1 [DataFields 集合 属性与方法](#41-datafields-集合-属性与方法)
   - 4.2 [DataField 对象属性](#42-datafield-对象属性)
   - 4.3 [cdrDataFieldType 枚举](#43-cdrdatafieldtype-枚举)
5. [DataItems 集合 — 形状上的数据视图](#五dataitems-集合--形状上的数据视图)
   - 5.1 [DataItems 属性与方法](#51-dataitems-属性与方法)
6. [DataItem 对象 — 单条数据项](#六dataitem-对象--单条数据项)
   - 6.1 [DataItem 属性一览](#61-dataitem-属性一览)
   - 6.2 [cdrDataItemType 枚举](#62-cdrdataitemtype-枚举)
7. [完整操作示例（VBA）](#七完整操作示例vba)
   - 7.1 [定义文档字段并为形状写入值](#71-定义文档字段并为形状写入值)
   - 7.2 [读取形状的对象数据](#72-读取形状的对象数据)
   - 7.3 [使用 Shape.Properties（自由键值对）](#73-使用-shapeproperties自由键值对)
   - 7.4 [公式驱动的数据项](#74-公式驱动的数据项)
   - 7.5 [批量遍历所有形状的对象数据](#75-批量遍历所有形状的对象数据)
   - 7.6 [C# 通过 COM 互操作读写对象数据](#76-c-通过-com-互操作读写对象数据)
8. [常见陷阱与注意事项](#八常见陷阱与注意事项)
9. [API 快速索引](#九api-快速索引)

---

## 一、概念背景与应用场景

CorelDRAW 的**对象数据**（Object Data）是一套附加在形状（Shape）上的**自定义元数据系统**，对应 UI 菜单路径：

```
Edit → Object Properties → Object Data（对象数据）
```

它允许为每个形状存储任意结构化的键值对，典型应用场景：

| 场景 | 说明 |
|------|------|
| 产品信息标注 | 为零件图形存储型号、材质、价格等业务字段 |
| 数据驱动排版 | 从 CSV/数据库导入数据，自动填充形状属性 |
| 跨形状批量统计 | 遍历页面所有形状，汇总某字段的值（如统计总价）|
| 脚本工作流标记 | 脚本处理过程中给形状打状态标签，方便后续过滤 |
| 与外部系统集成 | 用对象数据存储外部 ID，实现图形与数据库的双向关联 |

---

## 二、对象模型结构

```
Application
└── ActiveDocument (Document)
    ├── DataFields                   → DataFields 集合（文档级字段定义，相当于表头）
    │   ├── DataField("ProductName") → 字段定义（名称、类型、默认值）
    │   ├── DataField("Price")
    │   └── DataField("InStock")
    └── Pages
        └── Page
            └── Layer
                └── Shape
                    ├── Properties  → DataItems（自由键值对，无需预定义字段）
                    │   └── DataItem("注释") → 单条数据项（Name/Value/Type）
                    └── ObjectData  → DataItems（结构化数据，挂载到 DataFields 字段）
                        ├── DataItem("ProductName")
                        ├── DataItem("Price")
                        └── DataItem("InStock")
```

**关键路径：**

```vba
' 文档字段定义
Dim fields As DataFields
Set fields = ActiveDocument.DataFields

' 形状的自由属性
Dim props As DataItems
Set props = ActiveShape.Properties

' 形状的结构化对象数据
Dim items As DataItems
Set items = ActiveShape.ObjectData
```

---

## 三、Shape.Properties 与 Shape.ObjectData 的区别

| 对比项 | `Shape.Properties` | `Shape.ObjectData` |
|--------|-------------------|--------------------|
| 字段是否需要预定义 | ❌ 无需，可以直接 `Add` 任意名称 | ✅ 必须先在 `Document.DataFields` 中定义 |
| 是否显示在 UI 对象数据面板 | ❌ 不显示 | ✅ 显示 |
| 是否支持公式 | ❌ 不支持 | ✅ 支持（`DataItem.Formula`）|
| 适用场景 | 脚本内部临时标记、轻量附加数据 | 面向最终用户的结构化元数据 |
| 与 `DataField` 的关联 | 无 | 有；字段类型由 `DataField.Type` 控制 |

---

## 四、Document.DataFields — 文档级字段定义

`DataFields` 是整个文档的字段注册表，相当于数据库的**表结构定义**。所有形状的 `ObjectData` 共享这套字段定义。

### 4.1 DataFields 集合 属性与方法

| 成员 | 类型/签名 | 说明 |
|------|-----------|------|
| `Count` | `Long`（只读） | 已定义的字段总数 |
| `Item(index As Long)` | `DataField` | 按 1-based 索引获取字段 |
| `Item(name As String)` | `DataField` | 按名称获取字段（不存在则报错）|
| `Add(Name, Type, [DefaultValue])` | → `DataField` | 新增字段定义 |
| `Remove(index 或 name)` | `Sub` | 删除字段定义（**同时清除所有形状上该字段的值**）|

**示例：**

```vba
Dim fields As DataFields
Set fields = ActiveDocument.DataFields

' 遍历所有字段
Dim i As Integer
For i = 1 To fields.Count
    Debug.Print fields.Item(i).Name, fields.Item(i).Type
Next i

' 添加字段
Dim f As DataField
Set f = fields.Add("Price", cdrDataFieldReal, 0.0)

' 删除字段
fields.Remove "Price"        ' 按名称
fields.Remove 1              ' 按索引
```

### 4.2 DataField 对象属性

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Name` | `String` | 读写 | 字段名称，在文档内唯一 |
| `Type` | `cdrDataFieldType` | 只读 | 字段的数据类型（创建后不可更改）|
| `DefaultValue` | `Variant` | 读写 | 新形状尚未赋值时的默认显示值 |

### 4.3 cdrDataFieldType 枚举

| 常量 | 值 | 说明 |
|------|----|------|
| `cdrDataFieldText` | 0 | 文本字符串 |
| `cdrDataFieldInteger` | 1 | 整数 |
| `cdrDataFieldReal` | 2 | 浮点数（小数）|
| `cdrDataFieldDate` | 3 | 日期 |
| `cdrDataFieldTime` | 4 | 时间 |
| `cdrDataFieldBoolean` | 5 | 布尔值（True / False）|

---

## 五、DataItems 集合 — 形状上的数据视图

`Shape.ObjectData` 和 `Shape.Properties` 均返回 `DataItems` 集合，操作 API 完全相同。

### 5.1 DataItems 属性与方法

| 成员 | 类型/签名 | 说明 |
|------|-----------|------|
| `Count` | `Long`（只读） | 数据项数量 |
| `Item(index As Long)` | `DataItem` | 按 1-based 索引获取 |
| `Item(name As String)` | `DataItem` | 按名称获取（不存在则报错）|
| `Add(Name, [Type], [Value])` | → `DataItem` | 添加新项（**仅 Properties 支持**；ObjectData 的字段由文档定义，无需 Add）|
| `Remove(index 或 name)` | `Sub` | 删除指定数据项 |
| `Exists(name As String)` | `Boolean` | 检查指定名称是否存在（避免 Item 访问越界）|

---

## 六、DataItem 对象 — 单条数据项

`DataItem` 是存储在形状上的一条具体键值记录。

### 6.1 DataItem 属性一览

| 属性 | 类型 | 读写 | 说明 |
|------|------|------|------|
| `Name` | `String` | 只读 | 字段 / 属性名称 |
| `Type` | `cdrDataItemType` | 只读 | 当前值的数据类型 |
| `Value` | `Variant` | **读写** | 当前存储的值（最常用；若为公式驱动则为计算结果）|
| `StaticValue` | `Variant` | 读写 | 不含公式的静态值（公式存在时赋值此属性可覆盖公式）|
| `Formula` | `String` | 读写 | 公式表达式，如 `"=Price*Qty"`；设为空字符串则恢复静态值 |
| `DisplayString` | `String` | 只读 | 值格式化后的显示字符串（已处理类型、日期格式、单位等）|
| `IsFormula` | `Boolean` | 只读 | 当前值是否由公式驱动 |

### 6.2 cdrDataItemType 枚举

| 常量 | 值 | 说明 |
|------|----|------|
| `cdrDataItemText` | 0 | 字符串 |
| `cdrDataItemInteger` | 1 | 整数 |
| `cdrDataItemReal` | 2 | 浮点小数 |
| `cdrDataItemDate` | 3 | 日期（用 VBA `Date` 类型赋值）|
| `cdrDataItemTime` | 4 | 时间（用 VBA `Time` 类型赋值）|
| `cdrDataItemBoolean` | 5 | 布尔值 |

---

## 七、完整操作示例（VBA）

### 7.1 定义文档字段并为形状写入值

```vba
Sub SetupObjectData()
    Dim doc As Document
    Set doc = ActiveDocument

    '--- 步骤 1：在文档中定义字段（相当于建表结构）---
    Dim fields As DataFields
    Set fields = doc.DataFields

    ' 安全添加字段：先检查是否已存在
    Dim fName As DataField, fPrice As DataField, fInStock As DataField

    On Error Resume Next
    Set fName    = fields("ProductName")
    Set fPrice   = fields("Price")
    Set fInStock = fields("InStock")
    On Error GoTo 0

    If fName Is Nothing    Then Set fName    = fields.Add("ProductName", cdrDataFieldText,    "")
    If fPrice Is Nothing   Then Set fPrice   = fields.Add("Price",       cdrDataFieldReal,    0.0)
    If fInStock Is Nothing Then Set fInStock = fields.Add("InStock",     cdrDataFieldBoolean, True)

    '--- 步骤 2：为当前选中形状写入对象数据 ---
    Dim s As Shape
    Set s = ActiveShape

    Dim items As DataItems
    Set items = s.ObjectData

    items("ProductName").Value = "齿轮-A01"
    items("Price").Value       = 128.5
    items("InStock").Value     = True

    MsgBox "写入完成！ProductName=" & items("ProductName").Value
End Sub
```

### 7.2 读取形状的对象数据

```vba
Sub ReadObjectData()
    Dim s As Shape
    Set s = ActiveShape

    Dim items As DataItems
    Set items = s.ObjectData

    If items.Count = 0 Then
        MsgBox "该形状没有对象数据"
        Exit Sub
    End If

    Dim msg As String
    Dim i As Integer
    For i = 1 To items.Count
        Dim item As DataItem
        Set item = items.Item(i)
        ' 使用 DisplayString 获取格式化字符串（日期/布尔值更友好）
        msg = msg & item.Name & " (" & item.Type & ") = " & item.DisplayString & vbCrLf
    Next i

    MsgBox msg
End Sub
```

### 7.3 使用 Shape.Properties（自由键值对）

```vba
Sub UseShapeProperties()
    Dim s As Shape
    Set s = ActiveShape

    Dim props As DataItems
    Set props = s.Properties

    '--- 写入（Properties 可自由 Add，不受文档字段约束）---
    If Not props.Exists("脚本备注") Then
        props.Add "脚本备注", cdrDataItemText, ""
    End If
    props("脚本备注").Value = "已由批处理脚本处理，版本 v2.1"

    If Not props.Exists("处理时间戳") Then
        props.Add "处理时间戳", cdrDataItemText, ""
    End If
    props("处理时间戳").Value = CStr(Now)

    '--- 读取 ---
    Debug.Print "备注：" & props("脚本备注").Value
    Debug.Print "时间：" & props("处理时间戳").Value

    '--- 删除 ---
    ' props.Remove "脚本备注"

    MsgBox "Properties 写入成功：" & props("脚本备注").Value
End Sub
```

### 7.4 公式驱动的数据项

```vba
Sub UseFormula()
    '--- 前提：文档中已定义 Price、Qty、Total 三个 Real 类型字段 ---
    Dim doc As Document
    Set doc = ActiveDocument

    Dim fields As DataFields
    Set fields = doc.DataFields

    On Error Resume Next
    Dim fP As DataField, fQ As DataField, fT As DataField
    Set fP = fields("Price")
    Set fQ = fields("Qty")
    Set fT = fields("Total")
    On Error GoTo 0

    If fP Is Nothing Then Set fP = fields.Add("Price", cdrDataFieldReal,    0)
    If fQ Is Nothing Then Set fQ = fields.Add("Qty",   cdrDataFieldInteger, 1)
    If fT Is Nothing Then Set fT = fields.Add("Total", cdrDataFieldReal,    0)

    '--- 为形状赋值 ---
    Dim s As Shape
    Set s = ActiveShape

    s.ObjectData("Price").Value = 50.0
    s.ObjectData("Qty").Value   = 3

    '--- 用公式自动计算 Total ---
    ' 公式语法：引用本形状其他字段名，前缀 "="
    s.ObjectData("Total").Formula = "=Price*Qty"

    Debug.Print "Total 值：" & s.ObjectData("Total").Value      ' 150
    Debug.Print "是否公式：" & s.ObjectData("Total").IsFormula  ' True

    '--- 取消公式，恢复静态值 ---
    ' s.ObjectData("Total").Formula = ""
    ' s.ObjectData("Total").Value = 999

    MsgBox "Total = " & s.ObjectData("Total").DisplayString & _
           "（公式：" & s.ObjectData("Total").Formula & "）"
End Sub
```

### 7.5 批量遍历所有形状的对象数据

```vba
Sub BatchReadAllShapes()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim page As page
    Set page = doc.ActivePage

    Dim shapes As Shapes
    Set shapes = page.Shapes

    Dim report As String
    report = "页面共 " & shapes.Count & " 个形状" & vbCrLf & vbCrLf

    Dim i As Integer, j As Integer
    For i = 1 To shapes.Count
        Dim s As Shape
        Set s = shapes.Item(i)

        Dim items As DataItems
        Set items = s.ObjectData

        report = report & "▶ 形状[" & i & "] 名称=" & s.Name & _
                          " 类型=" & s.Type & vbCrLf

        If items.Count = 0 Then
            report = report & "    （无对象数据）" & vbCrLf
        Else
            For j = 1 To items.Count
                Dim it As DataItem
                Set it = items.Item(j)
                report = report & "    " & it.Name & " = " & it.DisplayString & vbCrLf
            Next j
        End If
    Next i

    ' 输出到调试窗口（数据量大时推荐）
    Debug.Print report

    MsgBox report
End Sub
```

### 7.6 C# 通过 COM 互操作读写对象数据

```csharp
using System;
using System.Runtime.InteropServices;

class CorelObjectDataDemo
{
    static void Main()
    {
        // 获取已运行的 CorelDRAW 实例
        dynamic app = Marshal.GetActiveObject("CorelDRAW.Application");
        dynamic doc = app.ActiveDocument;

        // ── 步骤 1：确保文档字段已定义 ──
        dynamic fields = doc.DataFields;
        bool hasPrice = false;
        for (int i = 1; i <= fields.Count; i++)
        {
            if (fields.Item(i).Name == "Price") { hasPrice = true; break; }
        }
        if (!hasPrice)
        {
            // cdrDataFieldReal = 2
            fields.Add("Price", 2, 0.0);
        }

        // ── 步骤 2：读写当前页所有形状的对象数据 ──
        dynamic page = doc.ActivePage;
        dynamic shapes = page.Shapes;

        for (int i = 1; i <= shapes.Count; i++)
        {
            dynamic shape = shapes.Item(i);
            dynamic items = shape.ObjectData;

            Console.Write($"Shape[{shape.Name}]: ");

            // 写入
            items["Price"].Value = 99.9 * i;

            // 读取
            for (int j = 1; j <= items.Count; j++)
            {
                dynamic item = items.Item(j);
                Console.Write($"{item.Name}={item.DisplayString}  ");
            }
            Console.WriteLine();
        }

        // ── 步骤 3：使用 Properties 写自由标签 ──
        dynamic activeShape = app.ActiveShape;
        dynamic props = activeShape.Properties;

        // Exists 检查
        bool exists = props.Exists("Tag");
        if (!exists)
        {
            // cdrDataItemText = 0
            props.Add("Tag", 0, "");
        }
        props["Tag"].Value = "processed";
        Console.WriteLine("Tag = " + props["Tag"].Value);
    }
}
```

---

## 八、常见陷阱与注意事项

| # | 问题 | 原因 | 解决方法 |
|---|------|------|---------|
| 1 | `items("字段名")` 报错"索引超出范围" | 该形状没有此字段的 DataItem（未赋过值）| 先用 `items.Exists("字段名")` 检查，或 `On Error Resume Next` |
| 2 | `Properties.Add` 与 `ObjectData` 的 Add 行为不同 | `ObjectData` 字段由文档 DataFields 统一管理，不能对 ObjectData 直接 Add 未定义字段 | 先 `doc.DataFields.Add(...)` 注册字段，再访问 `shape.ObjectData(...)` |
| 3 | `doc.DataFields.Remove("xx")` 后形状数据全部丢失 | 删除字段定义会同步清除所有形状上该字段的值，且**不可撤销** | 删除前先导出数据；生产环境慎用 |
| 4 | `IsFormula = True` 时直接赋值 `Value` 无效 | 公式优先级高于静态值，直接赋 Value 会被公式覆盖 | 先 `item.Formula = ""` 清除公式，再赋 `item.Value = ...` |
| 5 | `DisplayString` 与 `Value` 返回结果不同 | 日期/时间/布尔类型的 `Value` 是原始 VBA 类型，`DisplayString` 是格式化字符串 | 展示给用户用 `DisplayString`；程序内部计算用 `Value` |
| 6 | `Shape.Properties` 的数据在 UI 中不可见 | Properties 是脚本专用的隐藏属性，不会出现在对象数据面板 | 若需用户可见，改用 `ObjectData`（需先在 DataFields 注册）|
| 7 | 遍历 `ObjectData` 时 Count 为 0 | 文档中虽然有 DataFields 定义，但该形状从未被赋值过 | 判断 `Count = 0` 后跳过，或先赋默认值 |
| 8 | C# 中用字符串索引 `items["Price"]` 报错 | `DataItems` 的默认索引器在 COM 互操作中有时不支持字符串 | 改用 `items.Item("Price")` 显式调用 Item 方法 |

---

## 九、API 快速索引

### 属性入口

| 入口 | 返回类型 | 位置 | 说明 |
|------|----------|------|------|
| `Document.DataFields` | `DataFields` | Document 对象 | 文档级字段定义集合 |
| `Shape.ObjectData` | `DataItems` | Shape 对象 | 形状的结构化对象数据 |
| `Shape.Properties` | `DataItems` | Shape 对象 | 形状的自由键值属性 |

### DataFields 集合

| 成员 | 签名 | 说明 |
|------|------|------|
| `Count` | `Long` | 字段数量 |
| `Item` | `(Long 或 String) → DataField` | 获取字段定义 |
| `Add` | `(Name, Type, [Default]) → DataField` | 添加字段 |
| `Remove` | `(Long 或 String)` | 删除字段（**不可撤销**）|

### DataField 对象

| 属性 | 类型 | 说明 |
|------|------|------|
| `Name` | `String`（读写） | 字段名 |
| `Type` | `cdrDataFieldType`（只读） | 数据类型 |
| `DefaultValue` | `Variant`（读写） | 默认值 |

### DataItems 集合

| 成员 | 签名 | 说明 |
|------|------|------|
| `Count` | `Long` | 项数量 |
| `Item` | `(Long 或 String) → DataItem` | 获取数据项 |
| `Add` | `(Name, [Type], [Value]) → DataItem` | 添加项（仅 Properties）|
| `Remove` | `(Long 或 String)` | 删除项 |
| `Exists` | `(String) → Boolean` | 检查项是否存在 |

### DataItem 对象

| 属性 | 类型 | 说明 |
|------|------|------|
| `Name` | `String`（只读）| 名称 |
| `Type` | `cdrDataItemType`（只读）| 数据类型 |
| `Value` | `Variant`（读写）| 当前值（含公式结果）|
| `StaticValue` | `Variant`（读写）| 静态值（不含公式）|
| `Formula` | `String`（读写）| 公式字符串，空串表示无公式 |
| `DisplayString` | `String`（只读）| 格式化显示字符串 |
| `IsFormula` | `Boolean`（只读）| 是否公式驱动 |

### 枚举常量

**cdrDataFieldType / cdrDataItemType（值相同）**

| 常量 | 值 |
|------|----|
| `cdrDataFieldText` / `cdrDataItemText` | 0 |
| `cdrDataFieldInteger` / `cdrDataItemInteger` | 1 |
| `cdrDataFieldReal` / `cdrDataItemReal` | 2 |
| `cdrDataFieldDate` / `cdrDataItemDate` | 3 |
| `cdrDataFieldTime` / `cdrDataItemTime` | 4 |
| `cdrDataFieldBoolean` / `cdrDataItemBoolean` | 5 |

---

> **参考手册页码：** DataItems/DataItem 见第 2 册"集合与对象"章节；DataFields/DataField 见第 2 册"Document 属性"章节；枚举常量见第 6 册。
