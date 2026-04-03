' =============================================================================
' CorelDraw 脚本创建表格示例
' 适用版本：CorelDRAW X7 及以上
' 语言：VBA (Visual Basic for Applications)
'
' 核心 API（来自官方《CorelDRAW X7 脚本参考手册》）：
'   创建表格：Layer.CreateCustomShape("Table", Left, Top, Right, Bottom, Cols, Rows)
'   访问 TableShape 对象：Shape.Custom
'   访问单元格：TableShape.Cell(Column, Row)  ← 注意列在前、行在后
'   单元格文字：TableCell.TextShape.Text.Story
'   单元格填充：TableCell.Fill.ApplyUniformFill color
'   合并单元格：TableCellRange.Merge
'   单元格边框：TableBorders.SetBorders(cdrTableBorderXxx, IVGOutline)
'   添加行/列：TableShape.AddRow([RowIndex])  /  TableShape.AddColumn([ColIndex])
'   删除行/列：TableRow.Delete  /  TableColumn.Delete
'   设置行高：TableRow.SetHeight(Height, ResizeTable)
'   设置列宽：TableColumn.SetWidth(Width, ResizeTable)
'
' 使用方法：
'   在 CorelDRAW 中，打开「工具 > 宏 > 宏编辑器（Alt+F11）」，
'   新建模块后粘贴本代码，将光标置于 Main 过程内，按 F5 运行。
' =============================================================================

Option Explicit

' =============================================================================
' 示例一：产品信息表（带表头样式、斑马纹、列宽比例）
' =============================================================================
Sub Main()
    ' -------------------------------------------------------------------------
    ' 1. 确认活动文档；若无则新建
    ' -------------------------------------------------------------------------
    If Application.Documents.Count = 0 Then
        CreateDocument
    End If

    Dim doc   As Document
    Dim s     As Shape          ' 表格 Shape
    Dim ts    As Object         ' TableShape（通过 Shape.Custom 访问）

    Set doc = ActiveDocument

    ' -------------------------------------------------------------------------
    ' 2. 用 CreateCustomShape("Table", ...) 创建表格
    '
    '    参数含义（文档单位，默认毫米）：
    '      Left   = 距页面左边距  20 mm
    '      Top    = 距页面顶边距  20 mm
    '      Right  = Left + 宽度  = 140 mm
    '      Bottom = Top  + 高度  = 70 mm
    '      Columns = 4
    '      Rows    = 5（表头 1 行 + 数据 4 行）
    ' -------------------------------------------------------------------------
    Set s  = ActiveLayer.CreateCustomShape("Table", 20, 20, 140, 70, 4, 5)
    Set ts = s.Custom   ' ts 即 TableShape 对象

    ' -------------------------------------------------------------------------
    ' 3. 填充列标题（第 1 行）
    '    Cell(Column, Row) —— 列索引在前，行索引在后，均从 1 开始
    ' -------------------------------------------------------------------------
    ts.Cell(1, 1).TextShape.Text.Story = "产品编号"
    ts.Cell(2, 1).TextShape.Text.Story = "产品名称"
    ts.Cell(3, 1).TextShape.Text.Story = "单价（元）"
    ts.Cell(4, 1).TextShape.Text.Story = "库存（件）"

    ' -------------------------------------------------------------------------
    ' 4. 填充数据行（第 2-5 行）
    ' -------------------------------------------------------------------------
    Dim rowData(1 To 4, 1 To 4) As String
    rowData(1, 1) = "P-001" : rowData(1, 2) = "圆珠笔" : rowData(1, 3) = "2.50"  : rowData(1, 4) = "1200"
    rowData(2, 1) = "P-002" : rowData(2, 2) = "笔记本" : rowData(2, 3) = "15.00" : rowData(2, 4) = "560"
    rowData(3, 1) = "P-003" : rowData(3, 2) = "文件夹" : rowData(3, 3) = "8.80"  : rowData(3, 4) = "340"
    rowData(4, 1) = "P-004" : rowData(4, 2) = "订书机" : rowData(4, 3) = "32.00" : rowData(4, 4) = "88"

    Dim r As Integer, c As Integer
    For r = 1 To 4
        For c = 1 To 4
            ts.Cell(c, r + 1).TextShape.Text.Story = rowData(r, c)
        Next c
    Next r

    ' -------------------------------------------------------------------------
    ' 5. 设置表头样式（第 1 行）：深蓝背景、白色加粗文字、居中
    ' -------------------------------------------------------------------------
    Dim headerRange As Object   ' TableCellRange
    Set headerRange = ts.CellRange(1, 1, 4, 1)  ' CellRange(ColStart, RowStart, ColEnd, RowEnd)

    Dim blueColor As New Color
    blueColor.RGBAssign 31, 73, 125

    Dim fillObj As Fill
    Set fillObj = ActiveDocument.CreateFill("HeaderFill")
    fillObj.ApplyUniformFill blueColor
    headerRange.ApplyFill fillObj

    For c = 1 To 4
        With ts.Cell(c, 1).TextShape.Text.Story
            .Alignment = cdrCenterAlignment
            .Words.All.Size = 10
        End With
    Next c

    ' -------------------------------------------------------------------------
    ' 6. 数据行斑马纹：偶数行浅灰背景
    ' -------------------------------------------------------------------------
    Dim grayColor As New Color
    grayColor.RGBAssign 242, 242, 242

    Dim grayFill As Fill
    Set grayFill = ActiveDocument.CreateFill("GrayFill")
    grayFill.ApplyUniformFill grayColor

    For r = 2 To 5
        If r Mod 2 = 0 Then
            ts.CellRange(1, r, 4, r).ApplyFill grayFill
        End If
    Next r

    ' -------------------------------------------------------------------------
    ' 7. 数字列（第 3、4 列）数据行右对齐
    ' -------------------------------------------------------------------------
    For r = 2 To 5
        ts.Cell(3, r).TextShape.Text.Story.Alignment = cdrRightAlignment
        ts.Cell(4, r).TextShape.Text.Story.Alignment = cdrRightAlignment
    Next r

    ' -------------------------------------------------------------------------
    ' 8. 设置列宽（比例 1 : 2 : 1.5 : 1.5，总宽 120 mm）
    '    SetWidth(Width, ResizeTable) — Width 单位与文档单位一致（毫米）
    ' -------------------------------------------------------------------------
    ts.Columns(1).SetWidth 20,  False
    ts.Columns(2).SetWidth 40,  False
    ts.Columns(3).SetWidth 30,  False
    ts.Columns(4).SetWidth 30,  True    ' 最后一列 ResizeTable=True 使整体自适应

    ' -------------------------------------------------------------------------
    ' 9. 设置整体外框线宽（直接操作 Shape.Outline）
    ' -------------------------------------------------------------------------
    s.Outline.Width = 0.5

    ' -------------------------------------------------------------------------
    ' 10. 选中表格，返回焦点
    ' -------------------------------------------------------------------------
    s.Selected = True

    MsgBox "产品信息表创建完成！" & vbCrLf & _
           "5 行 × 4 列，位于页面左上角 (20, 20) mm", _
           vbInformation, "CorelDRAW 表格示例"
End Sub

' =============================================================================
' 示例二：2009 年 1 月日历
' （忠实复现官方《CorelDRAW X7 脚本参考手册》中的经典示例，
'  原文见 Layer.CreateCustomShape 方法说明页）
' =============================================================================
Sub CreateCalendar_January2009()
    If Application.Documents.Count = 0 Then
        CreateDocument
    End If

    Dim s1 As Shape

    ' 创建 7 列 × 7 行的表格（含标题行和星期行各 1 行，日期 5 行）
    ' 参数：Left=1, Top=10, Right=5, Bottom=7, Columns=7, Rows=6
    Set s1 = ActiveLayer.CreateCustomShape("Table", 1, 10, 5, 7, 7, 6)

    ' 第 1 行填入星期缩写
    s1.Custom.Cell(1, 1).TextShape.Text.Story = "Sun"
    s1.Custom.Cell(2, 1).TextShape.Text.Story = "Mon"
    s1.Custom.Cell(3, 1).TextShape.Text.Story = "Tue"
    s1.Custom.Cell(4, 1).TextShape.Text.Story = "Wed"
    s1.Custom.Cell(5, 1).TextShape.Text.Story = "Thu"
    s1.Custom.Cell(6, 1).TextShape.Text.Story = "Fri"
    s1.Custom.Cell(7, 1).TextShape.Text.Story = "Sat"

    ' 在第 1 行上方插入新行（作为月份标题行）
    s1.Custom.AddRow 1

    ' 合并新增标题行的全部 7 个单元格
    s1.Custom.Rows(1).Cells.All.Merge

    ' 填入月份标题并设置字号和居中对齐
    s1.Custom.Cell(1, 1).TextShape.Text.Story = "January"
    s1.Custom.Cell(1, 1).TextShape.Text.Story.Words.All.Size = 22
    s1.Custom.Cell(1, 1).TextShape.Text.Story.Alignment = cdrCenterAlignment

    ' 从第 13 个单元格（即第 3 行第 6 列，对应 1 月 1 日 = 星期四）开始填入日期
    ' Cells 集合按从左到右、从上到下顺序编号，第 1 行合并为 1 个单元格 → 编号 1
    ' 第 2 行（星期行）7 个单元格 → 编号 2~8，日期从第 9 个起（第 3 行星期日）
    ' 2009-01-01 是星期四，对应第 3 行第 5 列 → 单元格编号 = 9+4 = 13
    Dim i As Integer
    For i = 1 To 31
        s1.Custom.Cells(i + 12).TextShape.Text.Story = i
    Next i

    ' 合并第 3 行无日期的空白单元格（编号 9~12，对应星期日~星期三）并填灰色
    s1.Custom.Cells.Range(9, 10, 11, 12).Merge

    Dim grayFill As Fill
    Set grayFill = ActiveDocument.CreateFill("EmptyCellFill")
    grayFill.ApplyUniformFill CreateRGBColor(220, 220, 220)
    s1.Custom.Cells.Range(9, 10, 11, 12).ApplyFill grayFill

    ' 设置整张表格外框线宽
    s1.Outline.Width = 0.05

    ' 设置标题行（第 1 行）的内部边框
    s1.Custom.Rows(1).Cells.All.Borders.All.Width = 0.05

    ' 为 1 月 1 日所在单元格（编号 13）加绿色高亮边框
    s1.Custom.Cells(10).Borders.All.Width = 0.05
    s1.Custom.Cells.Range(10).Borders.All.Color.RGBAssign 0, 255, 0

    s1.Selected = True

    MsgBox "January 2009 日历创建完成！", vbInformation, "CorelDRAW 日历示例"
End Sub
