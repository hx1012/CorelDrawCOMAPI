' =============================================================================
' CorelDraw 脚本创建表格示例
' 适用版本：CorelDRAW X7 及以上
' 语言：VBA (Visual Basic for Applications)
' 使用方法：在 CorelDRAW 中，打开「工具 > 宏 > 宏编辑器」，
'           新建模块后粘贴本代码，按 F5 运行 Main 过程。
' =============================================================================

Option Explicit

' -----------------------------------------------------------------------------
' 主入口：创建一个带格式和内容的示例表格
' -----------------------------------------------------------------------------
Sub Main()
    Dim doc      As Document
    Dim layer    As layer
    Dim tblShape As Shape
    Dim tbl      As Table

    ' 行数、列数
    Const ROWS    As Integer = 5
    Const COLS    As Integer = 4

    ' 表格尺寸（单位：毫米）
    Const TBL_WIDTH  As Double = 120
    Const TBL_HEIGHT As Double = 80

    ' 表格左上角在页面中的位置（毫米，原点在页面左下角）
    Const POS_X As Double = 30
    Const POS_Y As Double = 170

    ' -------------------------------------------------------------------------
    ' 1. 确认活动文档；若无则新建
    ' -------------------------------------------------------------------------
    If Application.Documents.Count = 0 Then
        Set doc = Application.Documents.Add
    Else
        Set doc = Application.ActiveDocument
    End If

    Set layer = doc.ActivePage.ActiveLayer

    ' -------------------------------------------------------------------------
    ' 2. 在当前图层上创建表格对象
    '    CreateTable(rows, cols, width_mm, height_mm)
    ' -------------------------------------------------------------------------
    Set tblShape = layer.CreateTable(ROWS, COLS, _
                                     MillimetersToPoints(TBL_WIDTH), _
                                     MillimetersToPoints(TBL_HEIGHT))

    ' 将表格移动到指定位置（SetPosition 参数为点数，原点在页面左下角）
    tblShape.SetPosition MillimetersToPoints(POS_X), _
                         MillimetersToPoints(POS_Y)

    ' -------------------------------------------------------------------------
    ' 3. 获取 Table 对象，进行格式化与内容填充
    ' -------------------------------------------------------------------------
    Set tbl = tblShape.Table

    ' 调用子过程分别设置格式和内容
    Call FormatTableHeader(tbl)
    Call FillTableContent(tbl)
    Call SetColumnWidths(tbl, COLS, TBL_WIDTH)

    ' -------------------------------------------------------------------------
    ' 4. 选中表格，让用户直观看到结果
    ' -------------------------------------------------------------------------
    tblShape.Selected = True

    MsgBox "表格创建完成！" & vbCrLf & _
           "规格：" & ROWS & " 行 × " & COLS & " 列" & vbCrLf & _
           "尺寸：" & TBL_WIDTH & " mm × " & TBL_HEIGHT & " mm", _
           vbInformation, "CorelDRAW 表格示例"
End Sub

' -----------------------------------------------------------------------------
' 格式化表头（第 1 行）
' - 背景填充深蓝色
' - 文字白色、加粗、居中
' -----------------------------------------------------------------------------
Private Sub FormatTableHeader(tbl As Table)
    Dim col As Integer
    Dim cell As Cell

    For col = 1 To tbl.Columns
        Set cell = tbl.Cell(1, col)

        ' 背景色：深蓝 #1F497D
        cell.Fill.UniformColor.RGBAssign 31, 73, 125

        ' 文字颜色：白色
        cell.Text.Story.TextRange.Fill.UniformColor.RGBAssign 255, 255, 255

        ' 字体：加粗、10 号
        With cell.Text.Story.TextRange.Font
            .Bold = True
            .Size = 10
        End With

        ' 水平居中
        cell.Text.Story.TextRange.Alignment = cdrCenterAlignment
    Next col
End Sub

' -----------------------------------------------------------------------------
' 填充表格内容
' - 第 1 行：列标题
' - 第 2-5 行：数据示例
' -----------------------------------------------------------------------------
Private Sub FillTableContent(tbl As Table)
    ' ----- 列标题 -----
    Dim headers(1 To 4) As String
    headers(1) = "产品编号"
    headers(2) = "产品名称"
    headers(3) = "单价（元）"
    headers(4) = "库存（件）"

    Dim col As Integer
    For col = 1 To 4
        tbl.Cell(1, col).Text.Story.TextRange.Text = headers(col)
    Next col

    ' ----- 数据行 -----
    Dim data(2 To 5, 1 To 4) As String
    data(2, 1) = "P-001" : data(2, 2) = "圆珠笔"     : data(2, 3) = "2.50"  : data(2, 4) = "1200"
    data(3, 1) = "P-002" : data(3, 2) = "笔记本"     : data(3, 3) = "15.00" : data(3, 4) = "560"
    data(4, 1) = "P-003" : data(4, 2) = "文件夹"     : data(4, 3) = "8.80"  : data(4, 4) = "340"
    data(5, 1) = "P-004" : data(5, 2) = "订书机"     : data(5, 3) = "32.00" : data(5, 4) = "88"

    Dim row As Integer
    For row = 2 To 5
        For col = 1 To 4
            Dim cell As Cell
            Set cell = tbl.Cell(row, col)

            ' 填入文字
            cell.Text.Story.TextRange.Text = data(row, col)

            ' 奇偶行交替底色（斑马纹）
            If row Mod 2 = 0 Then
                ' 偶数行：浅灰 #F2F2F2
                cell.Fill.UniformColor.RGBAssign 242, 242, 242
            Else
                ' 奇数行：白色
                cell.Fill.UniformColor.RGBAssign 255, 255, 255
            End If

            ' 数字列（第 3、4 列）右对齐
            If col >= 3 Then
                cell.Text.Story.TextRange.Alignment = cdrRightAlignment
            End If

            ' 统一字号
            cell.Text.Story.TextRange.Font.Size = 9
        Next col
    Next row
End Sub

' -----------------------------------------------------------------------------
' 设置各列宽度（按比例分配总宽度）
' 列比例：1 : 2 : 1.5 : 1.5
' -----------------------------------------------------------------------------
Private Sub SetColumnWidths(tbl As Table, totalCols As Integer, totalWidthMM As Double)
    Dim ratios(1 To 4) As Double
    ratios(1) = 1
    ratios(2) = 2
    ratios(3) = 1.5
    ratios(4) = 1.5

    Dim ratioSum As Double
    ratioSum = 0
    Dim i As Integer
    For i = 1 To totalCols
        ratioSum = ratioSum + ratios(i)
    Next i

    For i = 1 To totalCols
        Dim widthMM As Double
        widthMM = totalWidthMM * ratios(i) / ratioSum
        tbl.Column(i).Width = MillimetersToPoints(widthMM)
    Next i
End Sub

' -----------------------------------------------------------------------------
' 辅助：将毫米转换为磅（CorelDRAW 内部单位）
' 1 mm = 2.834645669291339 points
' -----------------------------------------------------------------------------
Private Function MillimetersToPoints(mm As Double) As Double
    MillimetersToPoints = mm * 2.834645669291339
End Function
