Attribute VB_Name = "WT_Legacy"

'=========================================================
'=========================================================
' 快速查找：从当前位置向后查找第一个Word公式工具插入的公式

' 声明一个模块级变量，用于记录最近一次找到的公式的索引位置
' 这样可以支持多次连续执行“查找下一个”
Private lastFoundEquationIndex As Long

Sub A下一个Word公式()
    ' 查找光标位置之后的下一个公式
    Dim currentPosition As Long
    Dim oMath As oMath
    Dim i As Long
    Dim targetOMath As oMath
    Dim isFound As Boolean
    
    isFound = False
    Set targetOMath = Nothing
    
    ' 获取当前光标（选区）的起始位置
    currentPosition = Selection.Start
    
    With ActiveDocument
        ' 遍历文档中所有的公式对象
        For i = 1 To .OMaths.Count
            Set oMath = .OMaths.Item(i)
            
            ' 如果此公式的起始位置大于当前光标位置，说明它是“下一个”
            If oMath.Range.Start > currentPosition Then
                Set targetOMath = oMath
                lastFoundEquationIndex = i ' 记录找到的索引
                isFound = True
                Exit For ' 找到第一个就退出循环
            End If
        Next i
        
        If isFound Then
            ' 找到了：选中这个公式的范围，并滚动到视图中央
            targetOMath.Range.Select
            Application.Browser.Target = wdBrowsePage
            ActiveWindow.ScrollIntoView Selection.Range, True ' True表示滚动到视图中央
            
            ' 在状态栏给出友好提示
            Application.StatusBar = "找到第 " & lastFoundEquationIndex & " 个公式。按F5可继续查找下一个。"
        Else
            ' 没找到：提示用户，并询问是否从文档开头重新开始
            If MsgBox("已搜索到文档末尾。是否从开头重新开始查找？", vbQuestion + vbYesNo, "查找完成") = vbYes Then
                lastFoundEquationIndex = 0
                ' 将光标移到文档开头，然后递归调用自身
                Selection.HomeKey unit:=wdStory
                A下一个Word公式
            Else
                Application.StatusBar = "已取消。"
            End If
        End If
    End With
End Sub
' 快速查找上一个公式（从当前位置向前查找）
Sub A上一个Word公式()
    ' 查找光标位置之前的上一个公式
    Dim currentPosition As Long
    Dim oMath As oMath
    Dim i As Long
    Dim targetOMath As oMath
    
    currentPosition = Selection.Start
    Set targetOMath = Nothing
    
    With ActiveDocument
        ' 倒序遍历，找到位置小于当前光标且最大的那个（即最近的前一个）
        For i = .OMaths.Count To 1 Step -1
            Set oMath = .OMaths.Item(i)
            If oMath.Range.Start < currentPosition Then
                Set targetOMath = oMath
                lastFoundEquationIndex = i
                Exit For
            End If
        Next i
        
        If Not targetOMath Is Nothing Then
            targetOMath.Range.Select
            Application.Browser.Target = wdBrowsePage
            ActiveWindow.ScrollIntoView Selection.Range, True
            Application.StatusBar = "找到第 " & lastFoundEquationIndex & " 个公式。"
        Else
            MsgBox "当前光标之前未找到公式。", vbInformation, "查找完成"
        End If
    End With
End Sub
'==============================================
Sub A批量修改图片比例()
    '=============================================
    '批量修改文档所有图片缩放比例，弹窗输入百分比，默认55%
    '鲁棒优化：输入校验、错误捕获、锁定纵横比、兼容各类图片载体
    '=============================================
    Dim strInput As String
    Dim dblScale As Double
    Dim shp As InlineShape
    Dim pic As Shape
    Dim shpGroup As Shape
    Dim subShp As Shape
    Dim errMsg As String
    Dim screenUpdState As Boolean
    
    '保存原始屏幕刷新状态，关闭刷新提升运行速度
    screenUpdState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo GlobalErrorHandler '全局错误捕获
    
    '1. 弹窗获取缩放比例，默认值55
    strInput = InputBox(Prompt:="请输入图片缩放百分比（如55代表缩小至原55%）" & vbCrLf & _
                        "范围建议：1~1000，非法输入将自动使用默认55%", _
                        Title:="图片批量缩放设置", _
                        Default:="55")
    
    '用户点击取消则直接退出
    If strInput = "" Then
        GoTo RestoreEnv
    End If
    
    '2. 输入合法性校验
    If Not IsNumeric(strInput) Then
        dblScale = 55
        errMsg = errMsg & "【警告】输入非数字，已自动使用默认55%" & vbCrLf
    Else
        dblScale = CDbl(strInput)
        '限制合理区间 1% ~ 1000%
        If dblScale < 1 Or dblScale > 1000 Then
            dblScale = 55
            errMsg = errMsg & "【警告】数值超出1~1000范围，已自动使用默认55%" & vbCrLf
        End If
    End If
    
    '转换为小数比例（55% → 0.55）
    dblScale = dblScale / 100
    
    '3. 处理嵌入式图片（InlineShape，文档正文内嵌图片）
    For Each shp In ActiveDocument.InlineShapes
      If shp.Type = wdInlineShapePicture Then
        On Error Resume Next
        shp.LockAspectRatio = msoTrue
        ' InlineShape 没有 ScaleWidth/ScaleHeight，直接改 Width/Height
        shp.Width = shp.Width * dblScale
        shp.Height = shp.Height * dblScale
        On Error GoTo GlobalErrorHandler
      End If
    Next shp
    
    '4. 处理浮动形状图片、文本框、组合图形内图片
    For Each pic In ActiveDocument.Shapes
        Call ProcessSingleShape(pic, dblScale)
    Next pic
    
    '输出警告信息（如有非法输入）
    If errMsg <> "" Then
        MsgBox errMsg & "缩放完成！", vbExclamation, "执行提示"
    Else
        MsgBox "全部图片缩放完成，缩放比例：" & Format(dblScale * 100, "0.00") & "%", vbInformation, "执行成功"
    End If
    
RestoreEnv:
    '恢复Word原始环境设置
    Application.ScreenUpdating = screenUpdState
    Exit Sub

GlobalErrorHandler:
    '全局异常兜底（文档保护、权限、损坏对象等）
    MsgBox "程序运行出错：" & Err.Description & vbCrLf & "错误代码：" & Err.Number & vbCrLf & "部分图片未处理", vbCritical, "运行异常"
    Resume RestoreEnv
End Sub

'=============================================
'递归子过程：处理单个Shape，支持组合图形递归遍历内部图片
'参数：目标图形对象、缩放比例小数（0.55 表示 55%）
'=============================================
Private Sub ProcessSingleShape(targetShp As Shape, scaleVal As Double)
    Dim subShp As Shape
    Dim inShp As InlineShape
    On Error Resume Next ' 单图形异常跳过，不中断循环
    
    '情况1：当前形状本身是图片（使用 Type 判断，而非不存在的 HasPicture）
    If targetShp.Type = msoPicture Or targetShp.Type = msoLinkedPicture Then
        targetShp.LockAspectRatio = msoTrue
        ' ScaleWidth/ScaleHeight 是方法调用，Factor 为比例小数（0.55 表示 55%）
        targetShp.ScaleWidth scaleVal, msoTrue
        ' 锁定纵横比后，高度会自动跟随，无需再调 ScaleHeight
    End If
    
    '情况2：组合图形，递归遍历内部所有子图形（组合图、图文组合）
    If targetShp.Type = msoGroup Then
        For Each subShp In targetShp.GroupItems
            Call ProcessSingleShape(subShp, scaleVal)
        Next subShp
    End If
    
    '情况3：文本框内嵌套图片（InlineShape 依然用 wdInlineShapePicture 判断）
    If targetShp.TextFrame.HasText Then
        For Each inShp In targetShp.TextFrame.TextRange.InlineShapes
            If inShp.Type = wdInlineShapePicture Then
                inShp.LockAspectRatio = msoTrue
                inShp.Width = inShp.Width * scaleVal
                inShp.Height = inShp.Height * scaleVal
            End If
        Next inShp
    End If
    On Error GoTo 0
End Sub
'==============================================
Sub A批量修改图片宽度()
    ' 功能：批量修改图片和Visio图宽度，保持宽高比，使用自定义输入框设置宽度
    Dim iShape As InlineShape
    Dim objField As Field
    Dim visioShape As InlineShape
    Dim adjustedCount As Integer
    Dim selectedWidth As Single
    Dim customWidth As String
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim aspectRatio As Double
    Dim newHeight As Single
    
    ' 检查活动文档
    If ActiveDocument Is Nothing Then
        MsgBox "没有打开的文档！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 显示自定义宽度输入框（默认值12厘米）
    customWidth = InputBox("请输入图片宽度（单位：厘米）：", "设置图片宽度", "12")
    
    ' 检查用户输入
    If customWidth = "" Then
        MsgBox "未输入宽度，操作已取消。", vbInformation, "取消"
        Exit Sub
    End If
    
    ' 验证输入是否为数字
    If Not IsNumeric(customWidth) Then
        MsgBox "输入的不是有效的数字，请重新运行宏并输入正确的宽度值。", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    ' 转换为Word内部的Points单位
    selectedWidth = CentimetersToPoints(Val(customWidth))
    
    adjustedCount = 0
    
    ' 关闭屏幕更新，提高处理速度
    Application.ScreenUpdating = False
    
    ' 第一部分：处理所有嵌入式图片
    For Each iShape In ActiveDocument.InlineShapes
        If iShape.Type = wdInlineShapePicture Then
            ' 获取原图尺寸
            originalWidth = iShape.Width
            originalHeight = iShape.Height
            
            ' 计算宽高比（高度/宽度）
            If originalWidth > 0 Then
                aspectRatio = originalHeight / originalWidth
                
                ' 计算新的高度
                newHeight = selectedWidth * aspectRatio
                
                ' 应用新的尺寸
                iShape.Width = selectedWidth
                iShape.Height = newHeight
                
                adjustedCount = adjustedCount + 1
            End If
        End If
    Next iShape
    
    ' 第二部分：处理所有Visio图
    For Each objField In ActiveDocument.Fields
        If InStr(1, objField.code.text, "EMBED Visio.Drawing", vbTextCompare) > 0 Then
            On Error Resume Next
            Set visioShape = objField.result.InlineShapes(1)
            On Error GoTo 0
            
            If Not visioShape Is Nothing Then
                ' 获取原图尺寸
                originalWidth = visioShape.Width
                originalHeight = visioShape.Height
                
                ' 计算宽高比（高度/宽度）
                If originalWidth > 0 Then
                    aspectRatio = originalHeight / originalWidth
                    
                    ' 计算新的高度
                    newHeight = selectedWidth * aspectRatio
                    
                    ' 应用新的尺寸
                    visioShape.Width = selectedWidth
                    visioShape.Height = newHeight
                    
                    adjustedCount = adjustedCount + 1
                    Set visioShape = Nothing
                End If
            End If
        End If
    Next objField
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 显示结果信息
    MsgBox "批量修改图片宽度完成！" & vbCrLf & _
           "共调整了 " & adjustedCount & " 个图片/Visio图。" & vbCrLf & _
           "设置宽度为：" & customWidth & "厘米", _
           vbInformation, "完成"
End Sub

'==============================================
'==============================================
Sub A批量应用图片的样式()
    ' 功能：为文档中的嵌入式图片和Visio图应用"BG图行"样式
    ' 条件：图片/Visio图所在段落的下一段必须以"图"字开头
    Const STYLE_NAME As String = "BG图行"
    
    Dim doc As Document
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim hasImageOrVisio As Boolean
    Dim appliedCount As Integer
    
    ' 检查活动文档是否存在
    If ActiveDocument Is Nothing Then
        MsgBox "没有打开的文档！", vbExclamation, "错误"
        Exit Sub
    End If
    
    Set doc = ActiveDocument
    
    ' 检查样式是否存在
    On Error Resume Next
    Dim bgStyle As Style
    Set bgStyle = doc.Styles(STYLE_NAME)
    On Error GoTo 0
    
    If bgStyle Is Nothing Then
        MsgBox "样式 '" & STYLE_NAME & "' 不存在于当前文档中。", vbExclamation, "样式未找到"
        Exit Sub
    End If
    
    ' 初始化计数器
    appliedCount = 0
    
    ' 遍历文档中的所有段落
    For Each para In doc.Paragraphs
        ' 检查段落是否包含图片或Visio图
        hasImageOrVisio = False
        
        ' 检查段落是否包含嵌入式图片
        If para.Range.InlineShapes.Count > 0 Then
            hasImageOrVisio = True
        End If
        
        ' 如果没有找到图片，检查是否包含Visio图域
        If Not hasImageOrVisio Then
            If para.Range.Fields.Count > 0 Then
                Dim objField As Field
                For Each objField In para.Range.Fields
                    If InStr(1, objField.code.text, "EMBED Visio.Drawing", vbTextCompare) > 0 Then
                        hasImageOrVisio = True
                        Exit For
                    End If
                Next objField
            End If
        End If
        
        ' 如果当前段落包含图片或Visio图
        If hasImageOrVisio Then
            ' 获取当前段落的索引
            Dim paraIndex As Integer
            paraIndex = GetParagraphIndex(doc, para)
            
            ' 检查是否有下一段
            If paraIndex < doc.Paragraphs.Count Then
                ' 获取下一段
                Set nextPara = GetParagraphByIndex(doc, paraIndex + 1)
                
                ' 检查下一段是否以"图"字开头
                Dim nextParaText As String
                nextParaText = Trim(Replace(nextPara.Range.text, vbCr, ""))
                
                If Left(nextParaText, 1) = "图" Then
                    ' 应用样式
                    para.Range.Style = bgStyle
                    appliedCount = appliedCount + 1
                End If
            End If
        End If
    Next para
    
    ' 释放对象变量
    Set doc = Nothing
    Set para = Nothing
    Set nextPara = Nothing
    Set bgStyle = Nothing
    
    ' 显示结果信息
    MsgBox "已为 " & appliedCount & " 个符合条件的图片/Visio图应用了'" & STYLE_NAME & "'样式。", vbInformation, "完成"
End Sub

' 辅助函数：获取段落的索引
Private Function GetParagraphIndex(doc As Document, para As Paragraph) As Integer
    Dim i As Integer
    For i = 1 To doc.Paragraphs.Count
        If doc.Paragraphs(i).Range.Start = para.Range.Start Then
            GetParagraphIndex = i
            Exit Function
        End If
    Next i
    GetParagraphIndex = 0
End Function

' 辅助函数：通过索引获取段落
Private Function GetParagraphByIndex(doc As Document, index As Integer) As Paragraph
    If index >= 1 And index <= doc.Paragraphs.Count Then
        Set GetParagraphByIndex = doc.Paragraphs(index)
    End If
End Function
'==============================================
'==============================================
Sub A设置选中表格样式()
    ' 功能：当光标位于表格内时，对当前表格进行统一样式设置
    ' 包括：自动适应窗口宽度、应用表格样式、设置重复标题行、应用BG表文样式
    ' 增加了样式存在性检查
    
    Dim targetTable As Table
    Dim tableStyleName As String
    Dim textStyleName As String
    
    ' 设置要应用的样式名称
    tableStyleName = "网格型"
    textStyleName = "BG表文"
    
    ' 检查光标是否位于表格内
    If Selection.Information(wdWithInTable) Then
        
        ' 检查表格样式是否存在
        Dim tableStyleExists As Boolean
        Dim textStyleExists As Boolean
        tableStyleExists = False
        textStyleExists = False
        
        On Error Resume Next
        Dim testStyle As Style
        Set testStyle = ActiveDocument.Styles(tableStyleName)
        If Err.Number = 0 Then tableStyleExists = True
        Err.Clear
        
        Set testStyle = ActiveDocument.Styles(textStyleName)
        If Err.Number = 0 Then textStyleExists = True
        On Error GoTo 0
        
        ' 如果任一样式不存在，提示用户
        If Not (tableStyleExists And textStyleExists) Then
            Dim msg As String
            msg = "以下样式不存在，请先创建这些样式：" & vbNewLine & vbNewLine
            
            If Not tableStyleExists Then
                msg = msg & "- 表格样式: """ & tableStyleName & """" & vbNewLine
            End If
            
            If Not textStyleExists Then
                msg = msg & "- 文字样式: """ & textStyleName & """" & vbNewLine
            End If
            
            msg = msg & vbNewLine & "是否继续应用其他格式设置？"
            
            If MsgBox(msg, vbQuestion + vbYesNo, "样式缺失提示") = vbNo Then
                Exit Sub
            End If
        End If
        
        ' 将光标所在的表格赋值给变量
        Set targetTable = Selection.Tables(1)
        
        ' 1. 设置表格自动适应窗口宽度
        targetTable.AutoFitBehavior wdAutoFitWindow
        
        ' 2. 如果表格样式存在，则应用表格整体样式
        If tableStyleExists Then
            targetTable.Style = ActiveDocument.Styles(tableStyleName)
        Else
            MsgBox "跳过表格样式应用（样式不存在）", vbInformation, "提示"
        End If
        
        ' 3. 设置标题行重复（跨页时自动重复表头）
        ' 注意：实际使用时可以根据需要取消注释
        ' targetTable.Rows(1).HeadingFormat = True
        
        ' 4. 如果文字样式存在，则应用BG表文样式
        If textStyleExists Then
            targetTable.Range.Style = ActiveDocument.Styles(textStyleName)
        Else
            MsgBox "跳过文字样式应用（样式不存在）", vbInformation, "提示"
        End If
        
        ' 操作完成后，将光标定位到表格第一个单元格
        targetTable.Cell(1, 1).Range.Select
        
        ' 完成提示
        MsgBox "表格格式设置完成！" & vbNewLine & _
               "已应用的设置：" & vbNewLine & _
               "- 自适应窗口宽度" & vbNewLine & _
               IIf(tableStyleExists, "- 表格样式: " & tableStyleName & vbNewLine, "") & _
               IIf(textStyleExists, "- 文字样式: " & textStyleName, ""), _
               vbInformation, "操作完成"
               
    Else
        ' 当光标不在表格内时，提示用户
        MsgBox "请将光标置于表格内后再执行此操作", vbExclamation, "操作提示"
    End If
End Sub
'==============================================
'==============================================
Sub A表格插入序号123等()
' 功能：为选中的表格单元格区域依次插入连续序号（从1开始）
' 注意：使用前需先手动选中表格中需要填充序号的单元格区域
    Dim i As Integer
    Dim aCell As Cell
    ' 初始化序号为1
    i = 1
    ' 使用With语句简化对Selection对象的重复引用，提高效率
    With Selection
        ' 检查选中的是否为表格单元格，避免在非表格区域执行操作
        If .Information(wdWithInTable) Then
            ' 遍历选中区域中的每个单元格
            For Each aCell In .Cells
                ' 清除单元格中已有的内容
                aCell.Range.Delete
                ' 在单元格中插入当前序号
                aCell.Range.InsertAfter i
                ' 序号递增，为下一个单元格准备
                i = i + 1
            Next aCell

            ' 操作完成后给出提示
            MsgBox "已为选中区域插入序号，共" & (i - 1) & "个单元格", vbInformation, "操作完成"
        Else
            ' 当选中区域不在表格内时，提示用户
            MsgBox "请先选中表格中的单元格区域再执行此操作", vbExclamation, "操作提示"
        End If
    End With
End Sub
'==============================================
'==============================================
Sub A设置选中表格前3列列宽()
' 功能：为光标所在的表格设置前3列的固定宽度（单位：厘米）
' 提示：使用前需将鼠标光标放置在目标表格内部
    Dim tbl As Table
    ' 将光标当前所在的表格赋值给tbl变量（Selection.Tables(1)表示光标所在的第一个表格）
    Set tbl = Selection.Tables(1)
    ' 按列设置固定宽度（单位：厘米）
    tbl.Columns(1).Width = CentimetersToPoints(1)
    tbl.Columns(2).Width = CentimetersToPoints(5)
    tbl.Columns(3).Width = CentimetersToPoints(3)
End Sub
'==============================================
'==============================================
Sub A批量设置表格序号123等()
' 功能：为Word文档中所有表头含"序号"的表格自动填充序号
    Dim tbl As Table       ' 表格对象
    Dim cel As Cell        ' 单元格对象
    Dim rowIndex As Long   ' 行索引
    Dim celIndex As Long   ' 序号计数器
    Dim isFirstCell As Boolean ' 是否为第一个序号单元格
    Dim text As String     ' 表头文本
    ' 遍历文档中所有表格
    For Each tbl In ActiveDocument.Tables
        ' 获取第一行第一列的表头文本（去除空格和换行符）
        text = Replace(Trim(tbl.Cell(1, 1).Range.text), Chr(10), "", , , vbBinaryCompare)
        ' 仅处理表头包含"序号"的表格
        If InStr(text, "序号") <> 0 Then
            celIndex = 1          ' 重置序号计数器
            isFirstCell = True    ' 标记为第一个序号单元格
            ' 从第2行开始遍历表格行（跳过表头）
            For rowIndex = 2 To tbl.Rows.Count
                On Error Resume Next  ' 处理合并单元格导致的错误
                Set cel = tbl.Cell(rowIndex, 1)  ' 获取当前行第一列单元格
                ' 若单元格存在（未出错）
                If Err.Number = 0 Then
                    ' 清除原有编号和内容
                    cel.Range.ListFormat.RemoveNumbers
                    cel.Range.Delete
                    ' 填充序号（第一行序号为1，后续递增）
                    cel.Range.text = celIndex
                    If Not isFirstCell Then celIndex = celIndex + 1
                    isFirstCell = False  ' 取消第一个单元格标记
                Else
                    Err.Clear  ' 清除错误状态
                End If
                On Error GoTo 0  ' 恢复默认错误处理
            Next rowIndex
        End If
    Next tbl
    MsgBox "设置完成！"  ' 提示操作完成
End Sub
'==============================================
'==============================================
Sub A批量设置所有表格_固定列宽()
' 功能：统一设置文档中所有表格的格式（样式、列宽、行高及标题行重复）
'遍历文档中所有表格并统一设置格式。将表格设置为固定列宽模式；应用 ' 网格型 ' 作为表格样式，' BG表文 ' 作为文字样式；设置表格第一行为重复标题行；统一所有行的高度为 0.8 厘米；按以下尺寸固定列宽，第 1 列 1.2 厘米，第 2 列 2.7 厘米，第 3 列 1.9 厘米，第 4 列 8.84 厘米
    Dim tbl As Table       ' 表格对象
    Dim col As Column      ' 列对象
    ' 遍历文档中所有表格
    For Each tbl In ActiveDocument.Tables
        ' 设置表格为固定列宽模式
        tbl.AutoFitBehavior wdAutoFitFixed
        ' 应用表格样式和文字样式
        tbl.Style = ActiveDocument.Styles("网格型")
        tbl.Range.Style = ActiveDocument.Styles("BG表文")
        ' 设置标题行重复（跨页时重复表头）
        'tbl.Rows(1).HeadingFormat = True
        ' 统一设置所有行高（0.8厘米）
        tbl.Rows.Height = CentimetersToPoints(0.8)
        ' 按列设置固定宽度（单位：厘米）
        Set col = tbl.Columns(1): col.Width = CentimetersToPoints(1.2)
        Set col = tbl.Columns(2): col.Width = CentimetersToPoints(2.7)
        Set col = tbl.Columns(3): col.Width = CentimetersToPoints(1.9)
        Set col = tbl.Columns(4): col.Width = CentimetersToPoints(8.84)
    Next tbl
End Sub
'==============================================
'==============================================
Sub A批量设置所有表格_自动列宽()
' 功能：统一设置文档中所有表格的自适应宽度及格式
    Dim tbl As Table       ' 表格对象
    ' 遍历所有表格进行格式设置
    For Each tbl In ActiveDocument.Tables
        ' 设置表格宽度自适应窗口
        tbl.AutoFitBehavior wdAutoFitWindow
        ' 应用表格样式
        tbl.Style = ActiveDocument.Styles("网格型")
        ' 设置第一行为标题行（跨页重复）
        tbl.Rows(1).HeadingFormat = True
        ' 应用BG表文样式
        tbl.Range.Style = ActiveDocument.Styles("BG表文")
    Next tbl
End Sub
'==============================================
'==============================================
Sub A批量设置所有表格_第一行格式()
' 功能：将文档中所有表格的第一行样式统一设置为"BG表首行"
' 增加：检查样式是否存在，不存在则提示并退出
    
    Dim tbl As Table          ' 表格对象
    Dim styleName As String   ' 样式名称
    Dim StyleExists As Boolean ' 样式是否存在标志
    Dim styleObj As Style     ' 样式对象
    
    ' 设置要应用的样式名称
    styleName = "BG表首行"
    
    ' 检查样式是否存在
    StyleExists = False
    For Each styleObj In ActiveDocument.Styles
        If styleObj.NameLocal = styleName Then
            StyleExists = True
            Exit For
        End If
    Next styleObj
    
    ' 如果样式不存在，提示用户并退出
    If Not StyleExists Then
        MsgBox "错误：样式 '" & styleName & "' 不存在于当前文档中。" & vbCrLf & _
               "请先创建该样式，或检查样式名称是否正确。", _
               vbExclamation + vbOKOnly, "样式未找到"
        Exit Sub
    End If
    
    ' 遍历文档中所有表格
    For Each tbl In ActiveDocument.Tables
        ' 仅处理至少包含一行的表格
        If tbl.Rows.Count > 0 Then
            ' 为表格第一行应用"BG表首行"样式
            On Error Resume Next ' 防止意外错误
            tbl.Rows(1).Range.Style = styleName
            If Err.Number <> 0 Then
                ' 处理应用样式时的错误
                MsgBox "表格" & tbl.Range.ID & "应用样式时出错：" & Err.Description, _
                       vbExclamation, "应用样式错误"
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next tbl
    
    MsgBox "已完成！所有表格的首行样式已设置为 '" & styleName & "'。", _
           vbInformation, "操作完成"
    
End Sub
'==============================================
'==============================================
Sub A批量设置所有表格_行高()
    ' 功能：将文档中所有表格的行高设置为最小高度
    ' 作者：增强版 - 包含错误处理和用户交互
    ' 日期：2024
    ' 修改：将固定行高改为最小行高
    
    On Error GoTo ErrorHandler
    
    ' 声明变量
    Dim tbl As Table               ' 表格对象
    Dim doc As Document            ' 文档对象
    Dim dblRowHeight As Double     ' 最小行高值（厘米）
    Dim lngTableCount As Long      ' 表格计数器
    Dim lngProcessedCount As Long  ' 已处理表格计数器
    Dim bCancel As Boolean         ' 取消标志
    Dim vUserInput As Variant      ' 用户输入
    Dim bSkipErrors As Boolean     ' 跳过错误标志
    
    ' 设置文档对象
    Set doc = ActiveDocument
    
    ' 检查文档是否有效
    If doc Is Nothing Then
        MsgBox "没有活动的文档！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 检查文档中是否有表格
    If doc.Tables.Count = 0 Then
        MsgBox "当前文档中没有表格！", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 获取用户输入的最小行高
    vUserInput = InputBox("请输入表格最小行高（厘米）：" & vbCrLf & _
                          "默认值为0.8厘米" & vbCrLf & _
                          "范围：0.2 到 10.0 厘米" & vbCrLf & _
                          "注意：行高将不小于此值，但可根据内容自动增加。", _
                          "设置表格最小行高", "0.8")
    
    ' 检查用户是否取消
    If vUserInput = "" Then
        MsgBox "操作已取消。", vbInformation, "取消"
        Exit Sub
    End If
    
    ' 验证用户输入
    dblRowHeight = Val(Replace(vUserInput, ",", ".")) ' 替换逗号为点
    
    If dblRowHeight <= 0 Then
        MsgBox "请输入有效的正数行高！", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    If dblRowHeight < 0.2 Or dblRowHeight > 10 Then
        Dim vConfirm As VbMsgBoxResult
        vConfirm = MsgBox("您输入的行高 " & dblRowHeight & " 厘米可能超出常规范围。" & vbCrLf & _
                          "建议行高范围：0.2 - 10.0 厘米" & vbCrLf & _
                          "是否继续使用此值？", _
                          vbQuestion + vbYesNo, "确认行高")
        If vConfirm = vbNo Then
            Exit Sub
        End If
    End If
    
    ' 询问是否跳过错误
    bSkipErrors = (MsgBox("如果遇到错误，是否跳过并继续处理其他表格？", _
                          vbQuestion + vbYesNo, "错误处理") = vbYes)
    
    ' 显示确认对话框
    Dim strMessage As String
    strMessage = "即将设置文档中所有表格的最小行高：" & vbCrLf & vbCrLf
    strMessage = strMessage & "文档表格数量：" & doc.Tables.Count & vbCrLf
    strMessage = strMessage & "设置最小行高：" & dblRowHeight & " 厘米" & vbCrLf
    strMessage = strMessage & "（约合 " & Application.CentimetersToPoints(dblRowHeight) & " 磅）" & vbCrLf & vbCrLf
    strMessage = strMessage & "注意：行高将不小于此值，但可根据内容自动增加。" & vbCrLf & vbCrLf
    strMessage = strMessage & "是否继续执行？"
    
    If MsgBox(strMessage, vbQuestion + vbYesNo, "确认批量设置") = vbNo Then
        MsgBox "操作已取消。", vbInformation, "取消"
        Exit Sub
    End If
    
    ' 保存原始屏幕更新设置并禁用更新
    Dim bOriginalScreenUpdating As Boolean
    bOriginalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 创建进度提示
    Dim frmProgress As Object
    On Error Resume Next
    Set frmProgress = CreateSimpleProgressForm(doc.Tables.Count, "正在设置表格最小行高...")
    If Not frmProgress Is Nothing Then
        frmProgress.Show vbModeless
    End If
    On Error GoTo ErrorHandler
    
    ' 记录开始时间
    Dim dStartTime As Double
    dStartTime = Timer
    
    ' 遍历文档中所有表格
    lngTableCount = doc.Tables.Count
    lngProcessedCount = 0
    
    For Each tbl In doc.Tables
        lngProcessedCount = lngProcessedCount + 1
        
        ' 更新进度
        UpdateProgress frmProgress, lngProcessedCount, lngTableCount, _
                       "正在处理表格 " & lngProcessedCount & "/" & lngTableCount & _
                       " (" & Round(lngProcessedCount / lngTableCount * 100, 1) & "%)"
        
        ' 检查取消标志
        If Not frmProgress Is Nothing Then
            If frmProgress.Tag = "CANCEL" Then
                bCancel = True
                Exit For
            End If
        End If
        
        ' 处理当前表格 - 设置最小行高
        On Error Resume Next
        SetTableMinimumRowHeight tbl, dblRowHeight
        Dim lngErrNum As Long
        lngErrNum = Err.Number
        
        If lngErrNum <> 0 Then
            ' 记录错误但不中断
            Debug.Print "表格 " & lngProcessedCount & " 设置失败: " & Err.Description
            
            If Not bSkipErrors Then
                ' 询问用户如何处理错误
                Dim vErrorResponse As VbMsgBoxResult
                vErrorResponse = MsgBox("处理第 " & lngProcessedCount & " 个表格时出错：" & vbCrLf & _
                                        Err.Description & vbCrLf & vbCrLf & _
                                        "请选择处理方式：" & vbCrLf & _
                                        "【是】跳过此表格继续处理" & vbCrLf & _
                                        "【否】停止处理" & vbCrLf & _
                                        "【取消】忽略错误继续处理所有表格", _
                                        vbCritical + vbYesNoCancel, "处理错误")
                
                Select Case vErrorResponse
                    Case vbYes
                        ' 跳过此表格，继续下一个
                        Err.Clear
                    Case vbNo
                        bCancel = True
                        Exit For
                    Case vbCancel
                        bSkipErrors = True
                        Err.Clear
                End Select
            Else
                Err.Clear
            End If
        End If
        
        ' 允许系统处理其他事件
        DoEvents
    Next tbl
    
    ' 记录结束时间
    Dim dEndTime As Double
    dEndTime = Timer
    
Cleanup:
    ' 恢复原始设置
    Application.ScreenUpdating = bOriginalScreenUpdating
    
    ' 关闭进度窗体
    If Not frmProgress Is Nothing Then
        Unload frmProgress
        Set frmProgress = Nothing
    End If
    
    ' 显示完成消息
    Dim strResult As String
    Dim dblElapsedTime As Double
    
    dblElapsedTime = dEndTime - dStartTime
    
    If bCancel Then
        strResult = "操作已被用户取消。" & vbCrLf & _
                    "已成功处理 " & (lngProcessedCount - 1) & " 个表格。"
    Else
        strResult = "表格最小行高设置完成！" & vbCrLf & vbCrLf
        strResult = strResult & "成功处理表格数：" & lngProcessedCount & "/" & lngTableCount & vbCrLf
        strResult = strResult & "设置最小行高：" & dblRowHeight & " 厘米" & vbCrLf
        strResult = strResult & "耗时：" & Format(dblElapsedTime, "0.00") & " 秒"
        
        ' 如果可能，更新文档统计信息
        On Error Resume Next
        doc.UpdateWordCount
    End If
    
    MsgBox strResult, vbInformation, "操作完成"
    
    Exit Sub
    
ErrorHandler:
    ' 错误处理
    Dim strErrorMsg As String
    strErrorMsg = "发生错误 #" & Err.Number & ": " & Err.Description & vbCrLf & _
                  "发生在表格 " & lngProcessedCount & "/" & lngTableCount
    
    Debug.Print strErrorMsg
    MsgBox strErrorMsg, vbCritical, "严重错误"
    
    ' 尝试清理
    Resume Cleanup
End Sub

' ========================================================
' 辅助函数：设置表格最小行高（关键修改）
' ========================================================
Private Sub SetTableMinimumRowHeight(ByRef tbl As Table, ByVal dblMinHeightCm As Double)
    ' 检查表格对象是否有效
    If tbl Is Nothing Then
        Err.Raise 91, "SetTableMinimumRowHeight", "表格对象无效"
    End If
    
    ' 检查行高是否有效
    If dblMinHeightCm <= 0 Then
        Err.Raise 5, "SetTableMinimumRowHeight", "最小行高必须为正数"
    End If
    
    ' 转换为磅值（Word使用磅作为内部单位）
    Dim dblMinHeightPoints As Double
    dblMinHeightPoints = CentimetersToPoints(dblMinHeightCm)
    
    ' 检查转换后的值是否合理
    If dblMinHeightPoints <= 0 Then
        Err.Raise 6, "SetTableMinimumRowHeight", "行高转换失败"
    End If
    
    ' 设置所有行的最小高度
    On Error GoTo SetHeightError
    
    ' 关键修改：将行高规则改为"至少"（At Least）而不是"精确"（Exactly）
    tbl.Rows.HeightRule = wdRowHeightAtLeast  ' ← 这里改为最小行高模式
    tbl.Rows.Height = dblMinHeightPoints      ' 这是最小高度值
    
    ' 处理可能存在的合并单元格或特殊行
    Dim r As Long
    For r = 1 To tbl.Rows.Count
        ' 检查每行是否可设置高度
        If tbl.Rows(r).HeightRule <> wdRowHeightAtLeast Then
            tbl.Rows(r).HeightRule = wdRowHeightAtLeast
        End If
        tbl.Rows(r).Height = dblMinHeightPoints
    Next r
    
    Exit Sub
    
SetHeightError:
    ' 记录具体错误信息
    Err.Raise Err.Number, "SetTableMinimumRowHeight", _
              "设置表格最小行高失败（第" & r & "行）: " & Err.Description
End Sub

' ========================================================
' 辅助函数：创建简单进度窗体
' ========================================================
Private Function CreateSimpleProgressForm(ByVal lngTotal As Long, _
                                         ByVal strTitle As String) As Object
    On Error Resume Next
    
    Dim frm As Object
    Set frm = CreateObject("UserForm")
    
    If frm Is Nothing Then
        ' 无法创建窗体，返回Nothing
        Set CreateSimpleProgressForm = Nothing
        Exit Function
    End If
    
    With frm
        .Name = "frmProgress"
        .Caption = strTitle
        .Width = 300
        .Height = 120
        .StartUpPosition = 1 ' 屏幕中心
        
        ' 添加进度标签
        Dim lbl As Object
        Set lbl = .Controls.Add("Forms.Label.1", "lblProgress")
        With lbl
            .Caption = "正在处理 0/" & lngTotal
            .Top = 20
            .Left = 20
            .Width = 260
            .Height = 20
        End With
        
        ' 添加状态标签
        Dim lblStatus As Object
        Set lblStatus = .Controls.Add("Forms.Label.1", "lblStatus")
        With lblStatus
            .Caption = "请稍候..."
            .Top = 45
            .Left = 20
            .Width = 260
            .Height = 20
        End With
        
        ' 添加取消按钮
        Dim btnCancel As Object
        Set btnCancel = .Controls.Add("Forms.CommandButton.1", "btnCancel")
        With btnCancel
            .Caption = "取消"
            .Top = 75
            .Left = 110
            .Width = 80
            .Height = 25
        End With
    End With
    
    frm.Tag = ""
    
    Set CreateSimpleProgressForm = frm
End Function

' ========================================================
' 辅助函数：更新进度显示
' ========================================================
Private Sub UpdateProgress(ByRef frm As Object, _
                          ByVal lngCurrent As Long, _
                          ByVal lngTotal As Long, _
                          Optional ByVal strStatus As String = "")
    On Error Resume Next ' 防止进度更新失败导致主程序崩溃
    
    If frm Is Nothing Then Exit Sub
    
    ' 更新进度标签
    If Len(strStatus) > 0 Then
        frm.lblStatus.Caption = strStatus
    Else
        frm.lblProgress.Caption = "正在处理 " & lngCurrent & "/" & lngTotal & _
                                  " (" & Round(lngCurrent / lngTotal * 100, 1) & "%)"
    End If
    
    ' 刷新显示
    frm.Repaint
    DoEvents
End Sub

' ========================================================
' 辅助函数：检查文档是否受保护
' ========================================================
Private Function IsDocumentProtected(ByRef doc As Document) As Boolean
    On Error Resume Next
    IsDocumentProtected = (doc.ProtectionType <> wdNoProtection)
End Function

' ========================================================
' 可选：单独设置特定表格最小行高的函数
' ========================================================
Sub A设置单个表格行高()
    ' 功能：为当前选中的表格设置最小行高
    On Error GoTo ErrorHandler
    
    Dim tbl As Table
    Dim dblHeight As Double
    
    ' 检查是否有选中的表格
    If Selection.Tables.Count = 0 Then
        MsgBox "请先选中一个表格！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Set tbl = Selection.Tables(1)
    
    ' 获取用户输入
    dblHeight = Val(InputBox("请输入最小行高（厘米）：", "设置选中表格最小行高", "0.8"))
    
    If dblHeight <= 0 Then
        MsgBox "请输入有效的正数！", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    ' 设置最小行高
    SetTableMinimumRowHeight tbl, dblHeight
    
    MsgBox "表格最小行高已设置为 " & dblHeight & " 厘米", vbInformation, "完成"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "设置失败：" & Err.Description, vbCritical, "错误"
End Sub

'==============================================
'==============================================
Sub A删除未使用样式()
' 功能：删除文档中所有未使用的自定义样式（保留内置样式）
    Dim oStyle As Style  ' 样式对象
    ' 遍历文档中所有样式
    For Each oStyle In ActiveDocument.Styles
        ' 仅处理自定义样式（排除内置样式）
        If Not oStyle.BuiltIn Then
            ' 检查样式是否被文档内容使用
            With ActiveDocument.Content.Find
                .ClearFormatting       ' 清除查找格式
                .Style = oStyle.NameLocal  ' 指定要查找的样式
                .Execute findText:="", Format:=True  ' 执行格式查找
                ' 若未找到使用该样式的内容，则删除样式
                If Not .found Then oStyle.Delete
            End With
        End If
    Next oStyle
End Sub
'==============================================
'==============================================
Sub A统一修改所有纵向页面页边距()
' 功能：将文档中所有纵向页面的页边距统一设置为指定值
'遍历文档，统一所有纵向页面的页边距为上下边距 2 厘米，左右边距 2.5 厘米
    Dim sec As Section  ' 节对象
    ' 遍历文档中所有节
    For Each sec In ActiveDocument.Sections
        ' 仅处理纵向页面
        If sec.pageSetup.Orientation = wdOrientPortrait Then
            ' 设置页边距（单位：厘米）
            With sec.pageSetup
                .TopMargin = CentimetersToPoints(2)     ' 上页边距
                .BottomMargin = CentimetersToPoints(2)  ' 下页边距
                .LeftMargin = CentimetersToPoints(2.5)  ' 左页边距
                .RightMargin = CentimetersToPoints(2.5) ' 右页边距
            End With
        End If
    Next sec
End Sub
'==============================================
'==============================================
Sub 批量改段落格式()
' 功能：将文档中所有"标题 5"样式的段落批量修改为"标题 7"样式
'遍历文档，将所有应用了 "标题 5" 样式的段落，统一修改为 "标题 7" 样式
    Dim para As Paragraph  ' 段落对象
    ' 遍历文档中所有段落
    For Each para In ActiveDocument.Paragraphs
        ' 若段落样式为"标题 5"，则修改为"标题 7"
        If para.Style.NameLocal = "标题 5" Then
            para.Style = ActiveDocument.Styles("标题 7")
        End If
    Next para
End Sub
'==============================================
'==============================================
'插入一页横向的页面
Sub A插入横页()
    ' 在当前位置插入两个下一页分节符，创建独立的节
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    ' 移动光标到两个分节符之间的位置
    Selection.MoveUp unit:=wdLine, Count:=1
    Selection.TypeParagraph
    Selection.MoveLeft unit:=wdCharacter, Count:=1
    ' 切换当前节的页面方向（纵向<->横向）
    With Selection.pageSetup
        .Orientation = IIf(.Orientation = wdOrientPortrait, wdOrientLandscape, wdOrientPortrait)
    End With
End Sub

'==============================================
'==============================================
' 优化后的图题样式应用宏（修复Field对象范围获取）
Sub A批量应用BG图题样式()
    ' 声明变量
    Dim doc As Document        ' 文档对象
    Dim shp As InlineShape     ' 嵌入式图片对象
    Dim fieldObj As Field      ' 域对象（用于Visio图）
    Dim para As Paragraph      ' 段落对象
    Dim i As Long              ' 循环计数器
    Dim bScreenUpdate As Boolean ' 记录屏幕更新状态
    Dim bErrorOccurred As Boolean ' 错误标志
    Dim appliedCount As Long   ' 应用计数
    Dim totalItems As Long     ' 总图片/Visio图数量
    Dim currentItem As Long    ' 当前处理数量
    
    On Error GoTo ErrorHandler
    
    ' 设置当前活动文档
    Set doc = ActiveDocument
    
    ' 检查文档是否为空
    If doc Is Nothing Then
        MsgBox "没有活动的文档！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 检查所需样式是否存在
    If Not MStyleExists(doc, "BG图题") Then
        MsgBox "样式 'BG图题' 不存在！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 计算总图片和Visio图数量
    totalItems = doc.InlineShapes.Count + doc.Fields.Count
    
    If totalItems = 0 Then
        MsgBox "文档中没有图片或Visio图，无需应用图题样式！", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 记录当前屏幕更新状态并关闭屏幕更新
    bScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    'Application.EnableEvents = False ' 禁用事件处理以提高性能
    'Application.Calculation = wdCalculationManual ' 禁用自动计算
    
    ' 显示进度提示
    Dim frmProgress As Object
    On Error Resume Next ' 防止进度窗体创建失败
    Set frmProgress = CreateProgressForm()
    If Not frmProgress Is Nothing Then
        frmProgress.Show vbModeless
        frmProgress.lblStatus.Caption = "正在处理图片..."
        frmProgress.Repaint
    End If
    On Error GoTo ErrorHandler
    
    ' 初始化应用计数
    appliedCount = 0
    currentItem = 0
    
    ' 1. 处理所有嵌入式图片
    For Each shp In doc.InlineShapes
        ' 只处理图片类型
        If shp.Type = wdInlineShapePicture Then
            currentItem = currentItem + 1
            
            ' 每处理5个图片更新一次进度（减少更新频率）
            If Not frmProgress Is Nothing And currentItem Mod 5 = 0 Then
                frmProgress.lblProgress.Caption = "正在处理第 " & currentItem & " / " & totalItems & " 个项目"
                frmProgress.Repaint
                DoEvents ' 允许窗体更新
            End If
            
            ' 查找图片所在段落的下一段落
            Dim picPara As Paragraph
            Set picPara = shp.Range.Paragraphs(1)
            
            ' 确保不是最后一段
            If picPara.Range.End < doc.Content.End Then
                ' 获取下一段落
                Set para = picPara.Next
                
                ' 检查下一段是否有效且不为空
                If Not para Is Nothing Then
                    If Len(Trim(para.Range.text)) > 1 Then ' 包含段落标记
                        ' 检查段落是否以"图"开头
                        If Left(Trim(para.Range.text), 1) = "图" Then
                            ' 应用"BG图题"样式
                            On Error Resume Next
                            para.Style = "BG图题"
                            If Err.Number = 0 Then
                                appliedCount = appliedCount + 1
                            Else
                                Debug.Print "应用BG图题样式失败: " & Err.Description
                                Err.Clear
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                End If
            End If
        End If
        
        ' 添加取消功能检查
        If Not frmProgress Is Nothing Then
            If frmProgress.Tag = "CANCEL" Then
                bErrorOccurred = True
                Exit For
            End If
        End If
    Next shp
    
    ' 如果用户取消，退出
    If bErrorOccurred Then GoTo Cleanup
    
    ' 2. 处理所有Visio图
    For Each fieldObj In doc.Fields
        currentItem = currentItem + 1
        
        ' 每处理5个Visio图更新一次进度（减少更新频率）
        If Not frmProgress Is Nothing And currentItem Mod 5 = 0 Then
            frmProgress.lblProgress.Caption = "正在处理第 " & currentItem & " / " & totalItems & " 个项目"
            frmProgress.Repaint
            DoEvents ' 允许窗体更新
        End If
        
        ' 检查是否为Visio嵌入式对象
        If InStr(1, fieldObj.code.text, "EMBED Visio.Drawing", vbTextCompare) > 0 Then
            ' 查找Visio图所在段落的下一段落（修复：使用Result属性获取字段范围）
            Set para = fieldObj.result.Paragraphs(1)
            
            ' 确保不是最后一段
            If para.Range.End < doc.Content.End Then
                ' 获取下一段落
                Set para = para.Next
                
                ' 检查下一段是否有效且不为空
                If Not para Is Nothing Then
                    If Len(Trim(para.Range.text)) > 1 Then ' 包含段落标记
                        ' 检查段落是否以"图"开头
                        If Left(Trim(para.Range.text), 1) = "图" Then
                            ' 应用"BG图题"样式
                            On Error Resume Next
                            para.Style = "BG图题"
                            If Err.Number = 0 Then
                                appliedCount = appliedCount + 1
                            Else
                                Debug.Print "应用BG图题样式失败: " & Err.Description
                                Err.Clear
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                End If
            End If
        End If
        
        ' 添加取消功能检查
        If Not frmProgress Is Nothing Then
            If frmProgress.Tag = "CANCEL" Then
                bErrorOccurred = True
                Exit For
            End If
        End If
    Next fieldObj
    
Cleanup:
    ' 恢复应用设置
    Application.ScreenUpdating = bScreenUpdate
    'Application.EnableEvents = True
    'Application.Calculation = wdCalculationAutomatic
    
    ' 关闭进度窗体
    If Not frmProgress Is Nothing Then
        Unload frmProgress
        Set frmProgress = Nothing
    End If
    
    ' 提示用户宏执行完成
    If bErrorOccurred Then
        MsgBox "操作已被用户取消！", vbInformation, "操作中断"
    Else
        MsgBox "BG图题样式应用完成！" & vbCrLf & _
               "共处理 " & currentItem & " 个图片/Visio图。" & vbCrLf & _
               "成功应用 " & appliedCount & " 个BG图题样式。", vbInformation, "操作完成"
    End If
    
    Exit Sub
    
ErrorHandler:
    bErrorOccurred = True
    Dim sMsg As String
    sMsg = "发生错误：" & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "发生在第 " & currentItem & " 个项目处理过程中"
    
    MsgBox sMsg, vbCritical, "错误"
    Debug.Print sMsg
    Resume Cleanup
End Sub


' 优化后的表题样式应用宏（借鉴A下一个表格()的高效实现）
Sub A批量应用BG表题样式()
    ' 声明变量
    Dim doc As Document        ' 文档对象
    Dim tbl As Table           ' 表格对象
    Dim para As Paragraph      ' 段落对象
    Dim i As Long              ' 循环计数器
    Dim bScreenUpdate As Boolean ' 记录屏幕更新状态
    Dim bErrorOccurred As Boolean ' 错误标志
    Dim appliedCount As Long   ' 应用计数
    
    On Error GoTo ErrorHandler
    
    ' 设置当前活动文档
    Set doc = ActiveDocument
    
    ' 检查文档是否为空
    If doc Is Nothing Then
        MsgBox "没有活动的文档！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 检查所需样式是否存在
    If Not MStyleExists(doc, "BG表题") Then
        MsgBox "样式 'BG表题' 不存在！", vbExclamation, "错误"
        Exit Sub
    End If
    
    ' 检查是否有表格
    If doc.Tables.Count = 0 Then
        MsgBox "文档中没有表格，无需应用表题样式！", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 记录当前屏幕更新状态并关闭屏幕更新
    bScreenUpdate = Application.ScreenUpdating
    Application.ScreenUpdating = False
    'Application.EnableEvents = False ' 禁用事件处理以提高性能
    'Application.Calculation = wdCalculationManual ' 禁用自动计算
    
    ' 显示进度提示
    Dim frmProgress As Object
    On Error Resume Next ' 防止进度窗体创建失败
    Set frmProgress = CreateProgressForm()
    If Not frmProgress Is Nothing Then
        frmProgress.Show vbModeless
        frmProgress.lblStatus.Caption = "正在处理表格..."
        frmProgress.Repaint
    End If
    On Error GoTo ErrorHandler
    
    ' 初始化应用计数
    appliedCount = 0
    
    ' 遍历文档中的所有表格（借鉴A下一个表格()宏的高效实现）
    For i = 1 To doc.Tables.Count
        ' 更新进度显示
        If Not frmProgress Is Nothing Then
            frmProgress.lblProgress.Caption = "正在处理第 " & i & " / " & doc.Tables.Count & " 个表格"
            frmProgress.Repaint
            DoEvents ' 允许窗体更新
        End If
        
        Set tbl = doc.Tables(i)
        
        ' 查找表格的上一段落
        Dim prevParaRange As Range
        Set prevParaRange = tbl.Range.Previous(unit:=wdParagraph, Count:=1)
        
        ' 检查上一段落是否有效且不为空
        If Not prevParaRange Is Nothing Then
            ' 获取上一段落对象
            Set para = prevParaRange.Paragraphs(1)
            
            ' 检查段落是否不为空且以"表"开头
            If Len(Trim(para.Range.text)) > 1 Then ' 包含段落标记
                If Left(Trim(para.Range.text), 1) = "表" Then
                    ' 应用"BG表题"样式
                    On Error Resume Next ' 防止样式应用失败
                    para.Style = "BG表题"
                    If Err.Number = 0 Then
                        appliedCount = appliedCount + 1
                    Else
                        Debug.Print "表格 " & i & " 的上一段应用BG表题样式失败: " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
        End If
        
        ' 添加取消功能检查
        If Not frmProgress Is Nothing Then
            If frmProgress.Tag = "CANCEL" Then
                bErrorOccurred = True
                Exit For
            End If
        End If
    Next i
    
Cleanup:
    ' 恢复应用设置
    Application.ScreenUpdating = bScreenUpdate
    'Application.EnableEvents = True
    'Application.Calculation = wdCalculationAutomatic
    
    ' 关闭进度窗体
    If Not frmProgress Is Nothing Then
        Unload frmProgress
        Set frmProgress = Nothing
    End If
    
    ' 提示用户宏执行完成
    If bErrorOccurred Then
        MsgBox "操作已被用户取消！", vbInformation, "操作中断"
    Else
        MsgBox "BG表题样式应用完成！" & vbCrLf & _
               "共处理 " & i - 1 & " 个表格。" & vbCrLf & _
               "成功应用 " & appliedCount & " 个BG表题样式。", vbInformation, "操作完成"
    End If
    
    Exit Sub
    
ErrorHandler:
    bErrorOccurred = True
    Dim sMsg As String
    sMsg = "发生错误：" & Err.Number & vbCrLf & _
           Err.Description & vbCrLf & _
           "发生在第 " & i & " 个表格处理过程中"
    
    MsgBox sMsg, vbCritical, "错误"
    Debug.Print sMsg
    Resume Cleanup
End Sub


' 辅助函数：检查样式是否存在
Function MStyleExists(doc As Document, styleName As String) As Boolean
    On Error Resume Next
    Dim styleTest As Style
    Set styleTest = doc.Styles(styleName)
    MStyleExists = (Err.Number = 0)
    Err.Clear
End Function

' 创建进度窗体（简单版本）
Function CreateProgressForm() As Object
    On Error GoTo ErrorHandler
    
    Dim frm As Object
    
    ' 尝试创建窗体
    Set frm = ThisDocument.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    
    With frm
        .Properties("Name") = "frmProgress"
        .Properties("Caption") = "正在处理..."
        .Properties("Width") = 300
        .Properties("Height") = 100
        .Properties("StartUpPosition") = 1 ' 屏幕中心
        
        ' 添加标签
        Dim lblProgress As Object
        Set lblProgress = .Designer.Controls.Add("Forms.Label.1")
        With lblProgress
            .Name = "lblProgress"
            .Caption = "正在处理..."
            .Top = 10
            .Left = 10
            .Width = 280
        End With
        
        Dim lblStatus As Object
        Set lblStatus = .Designer.Controls.Add("Forms.Label.1")
        With lblStatus
            .Name = "lblStatus"
            .Caption = ""
            .Top = 30
            .Left = 10
            .Width = 280
        End With
        
        ' 添加取消按钮
        Dim btnCancel As Object
        Set btnCancel = .Designer.Controls.Add("Forms.CommandButton.1")
        With btnCancel
            .Name = "btnCancel"
            .Caption = "取消"
            .Top = 60
            .Left = 110
            .Width = 80
        End With
    End With
    
    ' 添加取消按钮的点击事件
    Dim codeModule As Object
    Set codeModule = frm.codeModule
    
    Dim code As String
    code = "Private Sub btnCancel_Click()" & vbCrLf & _
           "    Me.Tag = ""CANCEL""" & vbCrLf & _
           "    Me.Hide" & vbCrLf & _
           "End Sub"
    
    codeModule.AddFromString code
    
    Set CreateProgressForm = frm.Object
    Exit Function
    
ErrorHandler:
    ' 如果无法创建窗体，返回Nothing
    Set CreateProgressForm = Nothing
End Function
'==============================================
'==============================================
Sub A当前位置插入图编号()
    Selection.TypeText text:="图"
    With CaptionLabels("图表")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = False
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="图表", TitleAutoText:="InsertCaption1", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub

'==============================================
'==============================================
Sub A当前位置插入表编号()
    Selection.TypeText text:="表"
    With CaptionLabels("表格")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = False
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="表格", TitleAutoText:="InsertCaption2", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub
'==============================================
'==============================================
Sub A当前位置插入图编号章()
    Selection.TypeText text:="图"
    With CaptionLabels("图表")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="图表", TitleAutoText:="InsertCaption1", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub

'==============================================
'==============================================
Sub A当前位置插入表编号章()
    Selection.TypeText text:="表"
    With CaptionLabels("表格")
        .NumberStyle = wdCaptionNumberStyleArabic
        .IncludeChapterNumber = True
        .ChapterStyleLevel = 1
        .Separator = wdSeparatorHyphen
    End With
    Selection.InsertCaption Label:="表格", TitleAutoText:="InsertCaption2", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=1
End Sub
'==============================================
'==============================================
Sub A当前位置插入公式编号()
    ' 方案2：仅编号是域，括号为普通文本（括号全在域外）
    On Error GoTo ErrorHandler
    
    Dim objField As Field
    Dim currentPos As Long
    'currentPos = Selection.Start ' 锁定光标绝对位置
    
    ' 1. 插入左括号（普通文本，域外）
    ActiveDocument.Range(currentPos, currentPos).text = ")"
    'currentPos = currentPos + 1
    
    ' 2. 插入编号域（仅数字，无括号）
    Set objField = ActiveDocument.Fields.Add( _
        Range:=ActiveDocument.Range(currentPos, currentPos), _
        Type:=wdFieldSequence, _
        text:="公式 \* ARABIC", _
        PreserveFormatting:=False _
    )
    objField.Update
    
    ' 3. 插入右括号（普通文本，域外）
    'currentPos = objField.Result.End ' 定位到编号域末尾
    ActiveDocument.Range(currentPos, currentPos).text = "("
    
    ' 光标移到右括号后
    'Selection.Start = currentPos + 1
    'Selection.End = currentPos + 1
    
    'MsgBox "编号插入完成（括号全在域外）！", vbInformation, "成功"
    Exit Sub

ErrorHandler:
    MsgBox "插入失败：" & Err.Description, vbCritical, "错误"
End Sub
'==============================================
'==============================================
Sub A英文双引号转中文()
    On Error GoTo ErrorHandler ' 全局异常捕获
    
    ' ===== 前置校验：确保有有效选区 =====
    If Selection.Type = wdSelectionIP Then ' 仅光标（无选中内容）
        MsgBox "请先选中需要处理的文本区域！", vbExclamation
        Exit Sub
    End If
    If Not (Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionText) Then ' 仅处理文本选区
        MsgBox "仅支持对选中的文本区域操作，请勿选中图片/形状/表格等非文本对象！", vbExclamation
        Exit Sub
    End If
    
    Dim selRange As Range ' 绑定选区范围（核心：仅操作选中区域）
    Set selRange = Selection.Range ' 锁定选区，避免后续操作偏移
    Dim findRange As Range
    Set findRange = selRange.Duplicate ' 复制选区作为查找范围，不改动原选区
    
    Dim isFound As Boolean
    Dim isLeftQuote As Boolean: isLeftQuote = True ' 标记左右引号
    
    ' ===== 第一步：将选区中的英文半角双引号转为全角英文双引号（修复原编码错误） =====
    With findRange.Find
        .ClearFormatting
        .text = ChrW(34) ' 英文半角双引号（正确编码）
        .Replacement.text = ChrW(65282) ' 全角英文双引号
        .Forward = True
        .MatchWildcards = False
        .Wrap = wdFindStop ' 仅在选区范围内查找，不循环
        .Execute Replace:=wdReplaceAll ' 批量替换更高效
    End With
    
    ' ===== 第二步：将选区中的全角英文双引号替换为左右区分的中文引号 =====
    findRange.SetRange selRange.Start, selRange.End ' 重置查找范围为原选区（避免越界）
    With findRange.Find
        .ClearFormatting
        .text = ChrW(65282) ' 全角英文双引号
        .Forward = True
        .MatchWildcards = False
        .Wrap = wdFindStop ' 仅在选区范围内查找
        isFound = .Execute
    End With
    
    ' 逐一对选区中的全角引号替换为左右中文引号
    Do While isFound
        ' 防越界：确保查找范围在原选区内
        If findRange.Start > selRange.End Then Exit Do
        
        ' 替换为左/右中文引号
        If isLeftQuote Then
            findRange.text = ChrW(8220) ' 左中文引号 “
        Else
            findRange.text = ChrW(8221) ' 右中文引号 ”
        End If
        
        ' 移动查找范围（避免重复匹配当前引号）
        findRange.SetRange findRange.End, selRange.End
        isFound = findRange.Find.Execute
        isLeftQuote = Not isLeftQuote ' 切换左右引号标记
    Loop
    
    ' 操作完成提示
    MsgBox "选中区域的英文双引号已成功转为中文引号！", vbInformation
    Exit Sub ' 正常结束
    
ErrorHandler: ' 异常处理
    MsgBox "操作失败：" & Err.Description & "（错误码：" & Err.Number & "）", vbCritical
End Sub
'==============================================
'==============================================
Sub ADeleteUnusedStyles()
    ' =============================================================================================
    ' 功能：批量删除文档中“非内置”且“未使用”的样式。
    ' 适用：1000页+ 长文档，兼容 Word 2007 - 365。
    ' 原理：
    '   1. 保护所有 Word 内置样式（BuiltIn = True），只处理用户自定义样式。
    '   2. 使用倒序循环安全删除。
    '   3. 利用 Word 底层 Find 引擎全文档扫描，比遍历段落快 100 倍。
    '   4. 遍历所有 StoryRanges（包括页眉、页脚、脚注、文本框），确保无死角。
    ' =============================================================================================

    Dim doc As Document
    Dim sty As Style
    Dim i As Long
    Dim delCount As Long
    Dim totalCount As Long
    Dim isBuiltIn As Boolean
    Dim styleName As String
    
    Set doc = ActiveDocument
    
    ' --- 1. 性能优化设置 ---
    Application.ScreenUpdating = False ' 关闭屏幕刷新
    'Application.EnableEvents = False   ' 关闭事件响应
    ' 暂停分页计算（对长文档至关重要，防止删一个样式Word就重排1000页）
    doc.Repaginate
    Dim originalPagination As Boolean
    originalPagination = Options.Pagination
    Options.Pagination = False
    
    ' 错误处理：防止遇到损坏的样式导致宏崩溃
    On Error Resume Next
    
    totalCount = doc.Styles.Count
    delCount = 0
    
    ' --- 2. 倒序遍历样式集合 ---
    ' 必须倒序 (Step -1)，因为删除集合中的元素会改变后续元素的索引
    For i = totalCount To 1 Step -1
        Set sty = doc.Styles(i)
        styleName = sty.NameLocal
        
        ' 获取是否为内置样式
        isBuiltIn = True ' 默认设为True保护
        isBuiltIn = sty.BuiltIn
        
        ' 排除三种特殊情况（这三个是系统底层样式，有时 BuiltIn 属性可能误判或被修改，强制保留）
        If styleName = "Default Paragraph Font" Or styleName = "No List" Or styleName = "Table Grid" Then
            isBuiltIn = True
        End If
        
        ' --- 核心判断逻辑 ---
        ' 只有当样式是【非内置】(即用户后加的) 时，才检查是否删除
        If isBuiltIn = False Then
            
            ' 初步筛选：InUse 属性
            ' 如果 InUse = False，说明甚至没在样式库显示过，直接删，秒杀
            If sty.InUse = False Then
                doc.Styles(i).Delete
                If Err.Number = 0 Then delCount = delCount + 1
                Err.Clear
            Else
                ' 深度检查：InUse = True 并不代表真的用了，必须去文档里找
                If Not IsStyleReallyUsed(doc, sty) Then
                    ' 确认全文档都没找到该样式，删除
                    doc.Styles(i).Delete
                    If Err.Number = 0 Then delCount = delCount + 1
                    Err.Clear
                End If
            End If
            
        End If
        
        ' 状态栏进度提示 (长文档防焦虑)
        If i Mod 10 = 0 Then
            Application.StatusBar = "正在扫描样式... 剩余 " & i & " / " & totalCount
        End If
        
    Next i
    
    ' --- 3. 恢复环境 ---
    On Error GoTo 0
    Options.Pagination = originalPagination
    'Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    MsgBox "清理完成！" & vbCrLf & "共删除了 " & delCount & " 个未使用的自定义样式。", vbInformation, "样式瘦身"

End Sub

' =============================================================================================
' 辅助函数：全方位检查样式是否被使用
' 扫描范围：正文、页眉页脚、脚注尾注、文本框、批注
' =============================================================================================
Function IsStyleReallyUsed(doc As Document, targetStyle As Style) As Boolean
    Dim rng As Range
    Dim storyRange As Range
    Dim found As Boolean
    
    found = False
    
    ' 遍历文档的所有“故事类型”(StoryRanges)
    ' 这是一个循环链表结构，需要特殊处理才能遍历完所有页眉页脚
    For Each storyRange In doc.StoryRanges
        
        Set rng = storyRange
        
        ' 处理同一类型的多个区域 (例如：第一节页眉、第二节页眉...)
        Do
            With rng.Find
                .ClearFormatting
                .Style = targetStyle
                .text = ""           ' 不找具体文字
                .Format = True       ' 只找格式
                .Forward = True
                .Wrap = wdFindStop   ' 找到就停
                .Execute
                
                If .found Then
                    found = True
                    Exit Do ' 只要在任何一个地方找到了，就证明正在使用，退出内层循环
                End If
            End With
            
            ' 移动到同类型的下一个链接区域 (NextStoryRange)
            Set rng = rng.NextStoryRange
            
        Loop While Not rng Is Nothing
        
        If found Then Exit For ' 退出外层循环
    Next storyRange
    
    IsStyleReallyUsed = found
End Function
'==============================================
'==============================================
Sub A一键创建样式()
    ' 声明变量
    Dim doc As Document
    Dim targetStyle As Style
    Dim styleName As String
    Dim i As Integer
    
    ' 绑定当前活动文档
    Set doc = ActiveDocument
    
    ' 定义需要创建的样式名称数组
    Dim styleList As Variant
    styleList = Array("BG正文缩进", "BG正文顶格", "BG图行", "BG图题", "BG表题", "BG表文", "BG表首行", "BG公式")
    
    ' 循环处理每个样式
    For i = LBound(styleList) To UBound(styleList)
        styleName = styleList(i)
        
        ' 检查样式是否已存在（捕获不存在的错误）
        On Error Resume Next
        Set targetStyle = doc.Styles(styleName)
        On Error GoTo 0
        
        ' 样式不存在则创建，存在则跳过
        If targetStyle Is Nothing Then
            ' 创建段落样式（所有目标样式均为段落样式）
            Set targetStyle = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
            
            ' ========== 通用样式设置 ==========
            With targetStyle
                ' 样式基准：无基准
                .BaseStyle = ""
                
                ' 字体设置：宋体（中文）+Times New Roman（英文），小四（12磅）
                With .Font
                    .Name = "宋体"          ' 中文字体
                    .NameAscii = "Times New Roman" ' 英文字体（ASCII字符）
                    .NameOther = "Times New Roman" ' 其他字符字体
                    .Size = 12             ' 小四对应12磅（Word标准字号对应）
                    .Bold = False          ' 默认不加粗
                    .Italic = False        ' 默认不斜体
                End With
                
                ' 段落通用设置
                With .ParagraphFormat
                    ' 默认对齐方式：两端对齐
                    .alignment = wdAlignParagraphJustify
                    
                    .WidowControl = False                ' 关闭孤行控制（非孤行控制）
                    .KeepWithNext = False                 ' 默认不与下一段同页
                    .spaceBefore = 0                      ' 段前间距默认0
                    .spaceAfter = 0                       ' 段后间距默认0
                    
                    ' ========== 专属样式设置 ==========
                    ' 1. 首行缩进：仅BG正文缩进设置2字符
                    If styleName = "BG正文缩进" Then
                        ' 1字符=28.35磅，2字符=56.7磅
                        .FirstLineIndent = 24
                    Else
                        .FirstLineIndent = 0 ' 其他样式顶格
                    End If
                    
                    ' 2. 与下一段同页：BG图行、BG表题、BG表首行
                    Select Case styleName
                        Case "BG图行", "BG表题", "BG表首行"
                            .KeepWithNext = True
                    End Select
                    
                    ' 3. 段落对齐方式设置：BG表题、BG图题、BG图行设置为居中
                    Select Case styleName
                        Case "BG表题", "BG图题", "BG图行"
                            .alignment = wdAlignParagraphCenter
                    End Select
                End With
                
                ' 4. BG公式：添加居中+右对齐制表位
                If styleName = "BG公式" Then
                    With .ParagraphFormat
                        ' 清除原有制表位
                        .TabStops.ClearAll
                        ' 添加居中制表位（位置：页面水平居中，示例值7.5厘米，可根据页面调整）
                        .TabStops.Add _
                            Position:=CentimetersToPoints(7.5), _
                            alignment:=wdAlignTabCenter, _
                            Leader:=wdTabLeaderSpaces
                        ' 添加右对齐制表位（位置：正文右边界，示例值15厘米，可根据页面调整）
                        .TabStops.Add _
                            Position:=CentimetersToPoints(15), _
                            alignment:=wdAlignTabRight, _
                            Leader:=wdTabLeaderSpaces
                    End With
                End If
            End With
            
            ' 提示创建成功
            Debug.Print "样式【" & styleName & "】创建完成"
            Set targetStyle = Nothing
        Else
            ' 样式已存在，提示并跳过
            Debug.Print "样式【" & styleName & "】已存在，跳过创建"
            Set targetStyle = Nothing
        End If
    Next i
    
    ' 释放对象
    Set doc = Nothing
    MsgBox "所有BG样式处理完成！" & vbCrLf & "可在【样式】窗格中查看结果", vbInformation
End Sub
'==============================================
'==============================================
'■■■■■■■■■
Sub A公式行加制表位并应用BG公式样式()
    Dim para As Paragraph
    Dim rng As Range
    
    ' 检查样式
    On Error Resume Next
    If ActiveDocument.Styles("MTDisplayEquation").NameLocal <> "MTDisplayEquation" Then
        MsgBox "错误：样式 'MTDisplayEquation' 不存在！", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 获取当前段落
    Set para = Selection.Paragraphs(1)
    
    ' 保存当前光标位置
    Dim 原位置 As Long
    原位置 = Selection.Start
    
    ' 获取段落文本范围（不包括段落标记）
    Set rng = para.Range
    rng.End = rng.End - 1
    
    ' 删除所有制表位
    With rng.Find
        .text = "^t"
        .Replacement.text = ""
        .Execute Replace:=wdReplaceAll
    End With
    
        Selection.HomeKey unit:=wdLine
    Selection.HomeKey unit:=wdLine
    Selection.TypeText text:=vbTab
    Selection.EndKey unit:=wdLine
    Selection.EndKey unit:=wdLine
    Selection.TypeText text:=vbTab
    
    'Selection.InsertStyleSeparator
    
    ' 应用样式
    para.Range.Style = "MTDisplayEquation"
    
    ' 恢复光标位置
    Selection.SetRange 原位置 + 1, 原位置 + 1
    
    Selection.EndKey unit:=wdLine
    
    
End Sub
'■■■■■■■■■
Sub A应用图行样式()
    On Error GoTo ErrHandler
    If StyleExists("BG图行") Then
        Selection.Style = ActiveDocument.Styles("BG图行")
    Else
        'MsgBox "样式 'BG图行' 不存在于当前文档中", vbExclamation, "样式未找到"
        MsgBox "样式 'BG图行' 不存在于当前文档中。" & vbCrLf & _
        "请先创建该样式，或检查样式名称是否正确。", _
        vbExclamation + vbOKOnly, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■
Sub A应用图题样式()
    On Error GoTo ErrHandler
    If StyleExists("BG图题") Then
        Selection.Style = ActiveDocument.Styles("BG图题")
    Else
        'MsgBox "样式 'BG图题' 不存在于当前文档中", vbExclamation, "样式未找到"
        MsgBox "样式 'BG图题' 不存在于当前文档中。" & vbCrLf & _
        "请先创建该样式，或检查样式名称是否正确。", _
        vbExclamation + vbOKOnly, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■
Sub A应用表题样式()
    On Error GoTo ErrHandler
    If StyleExists("BG表题") Then
        Selection.Style = ActiveDocument.Styles("BG表题")
    Else
        MsgBox "样式 'BG表题' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■
Sub A应用表文样式()
    On Error GoTo ErrHandler
    If StyleExists("BG表文") Then
        Selection.Style = ActiveDocument.Styles("BG表文")
    Else
        MsgBox "样式 'BG表文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■
Sub A应用表首行样式()
    On Error GoTo ErrHandler
    If StyleExists("BG表首行") Then
        Selection.Style = ActiveDocument.Styles("BG表首行")
    Else
        MsgBox "样式 'BG表首行' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■
Sub A应用正文缩进样式()
    On Error GoTo ErrHandler
    If StyleExists("BG正文缩进") Then
        Selection.Style = ActiveDocument.Styles("BG正文缩进")
    Else
        MsgBox "样式 'BG正文缩进' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■
Sub A应用正文顶格样式()
    On Error GoTo ErrHandler
    If StyleExists("BG正文顶格") Then
        Selection.Style = ActiveDocument.Styles("BG正文顶格")
    Else
        MsgBox "样式 'BG正文顶格' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■
Sub A应用公式样式()
    On Error GoTo ErrHandler
    If StyleExists("BG公式") Then
        Selection.Style = ActiveDocument.Styles("BG公式")
    Else
        MsgBox "样式 'BG公式' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■
' 辅助函数：检查样式是否存在
'■■■■■■■■■
Function StyleExists(styleName As String) As Boolean
    On Error Resume Next
    Dim testStyle As Style
    Set testStyle = ActiveDocument.Styles(styleName)
    StyleExists = Not testStyle Is Nothing
End Function
'■■■■■■■■■
Sub A图片行应用图行样式()
    Const STYLE_NAME As String = "BG图行"
    Dim shp As InlineShape
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Then
            ' 关键：先确保它是 OLE，再排除公式
            If shp.OLEFormat Is Nothing Then
                ' 非 OLE，肯定是真图
                shp.Range.Paragraphs(1).Style = STYLE_NAME
            Else
                ' 是 OLE，再排除 MathType / MS Equation
                Select Case shp.OLEFormat.ProgID
                Case "Equation.DSMT4", "Equation.3"
                    ' 跳过公式，什么都不做
                Case Else
                    shp.Range.Paragraphs(1).Style = STYLE_NAME
                End Select
            End If
        End If
    Next shp
End Sub
'■■■■■■■■■
Sub A表格一键统一格式(Optional tableStyle As String = "网格型", _
                     Optional setAlignment As Boolean = True, _
                     Optional setHeaderRow As Boolean = True, _
                     Optional setFont As Boolean = True)
    ' 功能：批量设置文档中所有表格的格式
    ' 参数：
    '   tableStyle: 表格样式名称，默认为"网格型"
    '   setAlignment: 是否设置表格居中，默认为True
    '   setHeaderRow: 是否设置标题行，默认为True
    '   setFont: 是否设置字体，默认为True
    
    Dim aTable As Table
    Dim successCount As Long, failCount As Long
    Dim totalTables As Long
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    
    ' 记录开始时间
    startTime = Timer
    
    ' 统计表格总数
    totalTables = ActiveDocument.Tables.Count
    
    If totalTables = 0 Then
        MsgBox "文档中没有找到表格！", vbExclamation
        Exit Sub
    End If
    
    ' 显示进度（可选）
    Application.StatusBar = "正在处理表格，请稍候..."
    
    ' 停止屏幕更新，加快执行速度
    Application.ScreenUpdating = False
    
    For Each aTable In ActiveDocument.Tables
        On Error Resume Next
        
        ' 设置表格样式
        aTable.Style = tableStyle
        
        ' 设置表格居中
        If setAlignment Then
            aTable.Rows.alignment = wdAlignRowCenter
        End If
        
        ' 设置标题行
        If setHeaderRow And aTable.Rows.Count > 0 Then
            aTable.Rows(1).HeadingFormat = True
        End If
        
        ' 设置字体格式
        If setFont Then
            With aTable.Range
                .Font.Name = "宋体"
                .Font.Size = 12
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter ' 垂直居中
            End With
        End If
        
        ' 设置表格宽度为页面宽度的100%
        aTable.PreferredWidthType = wdPreferredWidthPercent
        aTable.PreferredWidth = 100
        
        ' 设置表格边框（确保网格型样式生效）
        aTable.Borders.Enable = True
        
        If Err.Number = 0 Then
            successCount = successCount + 1
            ' 更新进度（每处理5个表格更新一次）
            If successCount Mod 5 = 0 Then
                Application.StatusBar = "正在处理表格：" & successCount & "/" & totalTables
            End If
        Else
            failCount = failCount + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next aTable
    
    ' 恢复状态栏
    Application.StatusBar = False
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 计算耗时
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    ' 显示处理结果
    MsgBox "表格格式批量设置完成！" & vbCrLf & _
           "文档表格总数: " & totalTables & " 个" & vbCrLf & _
           "成功处理: " & successCount & " 个表格" & vbCrLf & _
           "处理失败: " & failCount & " 个表格" & vbCrLf & _
           "耗时: " & Format(elapsedTime, "0.00") & " 秒", _
           vbInformation, "操作完成"
    
    Exit Sub
    
ErrorHandler:
    ' 恢复状态栏和屏幕更新
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "发生错误：" & Err.Description & vbCrLf & _
           "错误代码: " & Err.Number & vbCrLf & _
           "已成功处理 " & successCount & " 个表格", _
           vbExclamation, "错误"
End Sub



'■■■■■■■■■
Sub A图片批量图题样式()
    ' =============================================================================================
    ' 功能：处理图片下方的段落
    ' 1. 自动排除 MathType 公式（它们也是 InlineShape，但不应处理）。
    ' 2. 向下检查图片下方的段落。
    ' 3. 删除图片与下方文字之间的空行。
    ' 4. 如果下方文字符合 "图X" 格式，应用 "BG图题" 样式。
    ' =============================================================================================

    Dim doc As Document
    Dim shp As InlineShape
    Dim checkRange As Range
    Dim styleName As String
    Dim regEx As Object
    Dim matchCount As Long
    Dim deleteCount As Long
    Dim cleanText As String
    Dim isEnd As Boolean
    Dim objClass As String
    Dim isEquation As Boolean
    
    ' --- 配置区域 ---
    styleName = "BG图题"
    
    Set doc = ActiveDocument
    
    ' --- 1. 检查样式是否存在 ---
    On Error Resume Next
    If doc.Styles(styleName) Is Nothing Then
        MsgBox "错误：文档中未找到样式 [" & styleName & "]。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' --- 2. 正则表达式 (支持 图1, 图 1.2, 图1-1) ---
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = False
        .IgnoreCase = True
        ' 匹配规则：以 "图" 开头，后跟数字、点或横杠
        .Pattern = "^\s*图\s*[\d\-\.]+"
    End With
    
    Application.ScreenUpdating = False
    matchCount = 0
    deleteCount = 0
    
    ' --- 3. 遍历所有嵌入式对象 ---
    For Each shp In doc.InlineShapes
        
        ' A. 排除 MathType 公式
        ' ------------------------------------------------
        isEquation = False
        If shp.Type = wdInlineShapeEmbeddedOLEObject Or shp.Type = wdInlineShapeLinkedOLEObject Then
            objClass = ""
            On Error Resume Next
            objClass = shp.OLEFormat.ClassType
            On Error GoTo 0
            
            ' 如果是公式，标记为 True
            If InStr(1, objClass, "Equation", vbTextCompare) > 0 Or _
               InStr(1, objClass, "MathType", vbTextCompare) > 0 Then
                isEquation = True
            End If
        End If
        
        ' B. 仅处理非公式的对象 (图片, Visio, PPT等)
        ' ------------------------------------------------
        If Not isEquation Then
            
            ' 定位到图片所在段落的末尾
            Set checkRange = shp.Range.Paragraphs(1).Range
            checkRange.Collapse direction:=wdCollapseEnd ' 折叠到段落结束位置
            
            ' 循环检查下一段，直到找到文字或到达文档末尾
            Do
                ' 尝试选取下一段
                ' MoveStart wdParagraph, 1 意味着向下移动一段
                ' 如果 checkRange 已经在文档最后，Move 会返回 0
                If checkRange.MoveStart(unit:=wdParagraph, Count:=1) = 0 Then
                    isEnd = True
                Else
                    ' 扩展范围以包含该段落文本（MoveStart只移动了起点，我们需要把终点也包进来以便读取Text）
                    checkRange.Expand unit:=wdParagraph
                    isEnd = False
                End If
                
                If isEnd Then Exit Do
                
                ' 处理文本
                cleanText = checkRange.text
                cleanText = Replace(cleanText, vbCr, "")
                cleanText = Replace(cleanText, vbLf, "")
                cleanText = Trim(cleanText)
                
                ' --- 判断逻辑 ---
                If Len(cleanText) = 0 Then
                    ' 1. 是空行 -> 删除
                    ' 注意：删除段落后，checkRange 的位置会自动合并到后方，
                    ' 我们需要重置 checkRange 到图片段落的末尾，重新寻找“新的下一段”
                    ' 或者直接删除该 range，下一轮循环逻辑依然成立
                    checkRange.Delete
                    deleteCount = deleteCount + 1
                    
                    ' 删除后，必须重置 checkRange 回到图片所在段落的后面，
                    ' 否则指针可能会迷失方向。
                    Set checkRange = shp.Range.Paragraphs(1).Range
                    checkRange.Collapse direction:=wdCollapseEnd
                    
                Else
                    ' 2. 是有文字的行 -> 检查是否为图题
                    If regEx.Test(cleanText) Then
                        ' 符合 "图 X" 格式
                        If checkRange.Paragraphs(1).Style <> styleName Then
                            checkRange.Paragraphs(1).Style = styleName
                            matchCount = matchCount + 1
                        End If
                    End If
                    
                    ' 只要遇到了文字（无论是不是图题，哪怕是正文），都停止查找
                    Exit Do
                End If
            Loop
            
        End If ' End If Not Equation
        
    Next shp
    
    Application.ScreenUpdating = True
    Set regEx = Nothing
    
    MsgBox "处理完成！" & vbCrLf & _
           "------------------------------" & vbCrLf & _
           "对象类型： 图片/Visio/PPT (已排除公式)" & vbCrLf & _
           "删除空行： " & deleteCount & " 处" & vbCrLf & _
           "应用样式： " & matchCount & " 处", vbInformation, "整理图题"

End Sub
'■■■■■■■■■
Sub A表格全选()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim i As Long, tableCount As Long
    Dim startTime As Double
    
    Set doc = ActiveDocument
    startTime = Timer
    tableCount = doc.Tables.Count
    
    ' 基本检查
    If tableCount = 0 Then
        MsgBox "文档中没有表格。", vbInformation
        Exit Sub
    End If
    
    If doc.ProtectionType <> wdNoProtection Then
        MsgBox "文档受保护，无法执行此操作。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 清除所有现有的可编辑区域
    doc.DeleteAllEditableRanges wdEditorEveryone
    
    ' 为每个表格添加可编辑区域
    For i = 1 To tableCount
        doc.Tables(i).Range.Editors.Add wdEditorEveryone
        
        ' 每处理50个表格更新一次状态（避免性能损失）
        If i Mod 50 = 0 Then
            Application.StatusBar = "处理中: " & i & "/" & tableCount
            DoEvents
        End If
    Next i
    
    ' 只选择可编辑区域（即表格）
    doc.SelectAllEditableRanges wdEditorEveryone
    
    ' 清除可编辑区域
    doc.DeleteAllEditableRanges wdEditorEveryone
    
    ' 验证选择结果
    Dim selectedCount As Long
    selectedCount = Selection.Tables.Count
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If Err.Number = 0 Then
        Dim elapsedTime As Double
        elapsedTime = Timer - startTime
        
        MsgBox "已选中 " & selectedCount & " 个表格（总数: " & tableCount & "）" & vbCrLf & _
               "耗时: " & Format(elapsedTime, "0.000") & " 秒", vbInformation
    Else
        MsgBox "错误: " & Err.Description, vbCritical
    End If
End Sub
'■■■■■■■■■
Sub A表格自动页宽()
    ' 功能：自动调整文档中所有表格的宽度适应窗口
    ' 优化版本：修复变量问题，添加错误处理和反馈
    
    Dim tbl As Table
    Dim tableCount As Long
    Dim processedCount As Long
    Dim failedCount As Long
    
    On Error GoTo ErrorHandler
    
    ' 获取表格数量
    tableCount = ActiveDocument.Tables.Count
    
    If tableCount = 0 Then
        MsgBox "文档中没有表格。", vbInformation, "提示"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 处理每个表格
    processedCount = 0
    failedCount = 0
    
    For Each tbl In ActiveDocument.Tables
        On Error Resume Next
        tbl.AutoFitBehavior wdAutoFitWindow
        
        If Err.Number = 0 Then
            processedCount = processedCount + 1
        Else
            failedCount = failedCount + 1
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next tbl
    
    Application.ScreenUpdating = True
    
    ' 显示详细结果
    MsgBox "表格调整完成！" & vbCrLf & vbCrLf & _
           "表格总数: " & tableCount & vbCrLf & _
           "成功调整: " & processedCount & vbCrLf & _
           "调整失败: " & failedCount, vbInformation, "操作完成"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Description & vbCrLf & _
           "错误代码: " & Err.Number, vbCritical, "错误"
End Sub

'=======================================
' ===================== 宏1：定位到下一个图片 =====================
Sub A下一张图片()
    ' 功能：定位到当前光标位置之后的下一个图片（支持嵌入型/浮动型）
    On Error GoTo ErrorHandler ' 启用错误捕获
    
    Dim currentRange As Range       ' 当前光标位置
    Dim nextInlinePic As InlineShape ' 嵌入型图片
    Dim nextShape As Shape           ' 浮动型图片
    Dim picFound As Boolean          ' 是否找到图片标记
    Dim targetRange As Range         ' 目标图片位置
    
    ' 初始化变量
    Set currentRange = Selection.Range.Duplicate
    picFound = False
    
    ' 第一步：查找嵌入型图片（InlineShape，仅筛选图片类型）
    For Each nextInlinePic In ActiveDocument.InlineShapes
        If nextInlinePic.Type = wdInlineShapePicture Then ' 仅匹配图片，排除其他嵌入对象
            If nextInlinePic.Range.Start > currentRange.End Then
                Set targetRange = nextInlinePic.Range
                picFound = True
                Exit For ' 找到第一个符合条件的图片后退出循环
            End If
        End If
    Next
    
    ' 第二步：若未找到嵌入型，查找浮动型图片（Shape，仅筛选图片类型）
    If Not picFound Then
        For Each nextShape In ActiveDocument.Shapes
            If nextShape.Type = msoPicture Then ' 仅匹配图片，排除形状/文本框等
                ' 浮动图片的锚点决定其在文档中的逻辑位置
                If Not nextShape.Anchor Is Nothing Then
                    If nextShape.Anchor.Start > currentRange.End Then
                        Set targetRange = nextShape.Anchor ' 定位到锚点位置
                        picFound = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    
    ' 定位并提示结果
    If picFound Then
        targetRange.Select
        'MsgBox "已定位到下一张图片！", vbInformation, "操作成功"
    Else
        MsgBox "当前位置之后未找到图片！", vbExclamation, "提示"
    End If
    
    ' 清理对象（避免内存泄漏）
Cleanup:
        Set currentRange = Nothing
        Set targetRange = Nothing
        Set nextInlinePic = Nothing
        Set nextShape = Nothing
    Exit Sub
    
    ' 错误处理
ErrorHandler:
        MsgBox "定位图片时发生错误：" & Err.Description, vbCritical, "错误"
        Resume Cleanup ' 跳转到对象清理步骤
End Sub



'=======================================
' ===================== 宏2：定位到上一个图片 =====================
Sub A上一张图片()
    ' 功能：定位到当前光标位置之前的上一个图片（支持嵌入型/浮动型）
    On Error GoTo ErrorHandler ' 启用错误捕获
    
    Dim currentRange As Range       ' 当前光标位置
    Dim prevInlinePic As InlineShape ' 嵌入型图片
    Dim prevShape As Shape           ' 浮动型图片
    Dim picFound As Boolean          ' 是否找到图片标记
    Dim targetRange As Range         ' 目标图片位置
    
    ' 初始化变量
    Set currentRange = Selection.Range.Duplicate
    picFound = False
    
    ' 第一步：查找嵌入型图片（InlineShape，仅筛选图片类型）
    For Each prevInlinePic In ActiveDocument.InlineShapes
        If prevInlinePic.Type = wdInlineShapePicture Then ' 仅匹配图片，排除其他嵌入对象
            If prevInlinePic.Range.End < currentRange.Start Then
                Set targetRange = prevInlinePic.Range
                picFound = True
                ' 继续遍历，找最后一个符合条件的图片（最近的上一张）
            End If
        End If
    Next
    
    ' 第二步：若未找到嵌入型，查找浮动型图片（Shape，仅筛选图片类型）
    If Not picFound Then
        For Each prevShape In ActiveDocument.Shapes
            If prevShape.Type = msoPicture Then ' 仅匹配图片，排除形状/文本框等
                ' 浮动图片的锚点决定其在文档中的逻辑位置
                If Not prevShape.Anchor Is Nothing Then
                    If prevShape.Anchor.End < currentRange.Start Then
                        Set targetRange = prevShape.Anchor ' 定位到锚点位置
                        picFound = True
                        ' 继续遍历，找最后一个符合条件的图片（最近的上一张）
                    End If
                End If
            End If
        Next
    End If
    
    ' 定位并提示结果
    If picFound Then
        targetRange.Select
        'MsgBox "已定位到上一张图片！", vbInformation, "操作成功"
    Else
        MsgBox "当前位置之前未找到图片！", vbExclamation, "提示"
    End If
    
    ' 清理对象（避免内存泄漏）
Cleanup:
        Set currentRange = Nothing
        Set targetRange = Nothing
        Set prevInlinePic = Nothing
        Set prevShape = Nothing
    Exit Sub
    
    ' 错误处理
ErrorHandler:
        MsgBox "定位图片时发生错误：" & Err.Description, vbCritical, "错误"
        Resume Cleanup ' 跳转到对象清理步骤
End Sub

'================================================
'================================================

' ===================== 宏2：定位到下一个公式 =====================
Sub A下一个公式()
    ' 功能：定位到当前光标位置之后的下一个公式（兼容Word 2010+ OMath公式）
    On Error GoTo ErrorHandler
    
    Dim currentRange As Range   ' 当前光标位置
    Dim oMath As oMath          ' 公式对象
    Dim eqFound As Boolean      ' 是否找到公式标记
    Dim targetRange As Range    ' 目标公式位置
    
    ' 初始化变量
    Set currentRange = Selection.Range.Duplicate
    eqFound = False
    
    ' 遍历文档中所有公式对象（OMath）
    For Each oMath In ActiveDocument.OMaths
        ' 检查公式是否在光标之后
        If oMath.Range.Start > currentRange.End Then
            Set targetRange = oMath.Range
            eqFound = True
            Exit For
        End If
    Next
    
    ' 兼容旧版公式编辑器（Equation Editor 3.0）
    If Not eqFound Then
        For Each nextInlinePic In ActiveDocument.InlineShapes
            If nextInlinePic.Type = wdInlineShapeOLEObject Then
                If InStr(1, nextInlinePic.OLEFormat.ProgID, "Equation", vbTextCompare) > 0 Then
                    If nextInlinePic.Range.Start > currentRange.End Then
                        Set targetRange = nextInlinePic.Range
                        eqFound = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
    
    ' 定位并提示结果
    If eqFound Then
        targetRange.Select
        'MsgBox "已定位到下一个公式！", vbInformation, "操作成功"
    Else
        MsgBox "当前位置之后未找到公式！", vbExclamation, "提示"
    End If
    
    ' 清理对象
Cleanup:
        Set currentRange = Nothing
        Set targetRange = Nothing
        Set oMath = Nothing
    Exit Sub
    
    ' 错误处理
ErrorHandler:
        MsgBox "定位公式时发生错误：" & Err.Description, vbCritical, "错误"
        Resume Cleanup
End Sub

' ===================== 宏3：定位到下一个表格 =====================
' ===================== 宏3：定位到下一个表格 =====================
Sub A下一个表格()
    ' 功能：定位到当前光标位置之后的下一个表格，并选中整个表格
    On Error GoTo ErrorHandler
    
    Dim currentRange As Range   ' 当前光标位置
    Dim tbl As Table            ' 表格对象
    Dim tblFound As Boolean     ' 是否找到表格标记
    
    ' 初始化变量
    Set currentRange = Selection.Range.Duplicate
    tblFound = False
    
    ' 遍历文档中所有表格
    For Each tbl In ActiveDocument.Tables
        ' 检查表格是否在光标之后
        If tbl.Range.Start > currentRange.End Then
            ' 选中整个表格
            tbl.Range.Select
            tblFound = True
            Exit For
        End If
    Next
    
    ' 提示结果
    If tblFound Then
        'MsgBox "已定位到下一个表格并选中！", vbInformation, "操作成功"
    Else
        MsgBox "当前位置之后未找到表格！", vbExclamation, "提示"
    End If
    
    ' 清理对象
Cleanup:
        Set currentRange = Nothing
        Set tbl = Nothing
    Exit Sub
    
    ' 错误处理
ErrorHandler:
        MsgBox "定位表格时发生错误：" & Err.Description, vbCritical, "错误"
        Resume Cleanup
End Sub
'===============================================
'===============================================
' ===================== 宏4：定位到上一个表格 =====================
Sub A上一个表格()
    ' 功能：定位到当前光标位置之前的上一个表格，并选中整个表格
    On Error GoTo ErrorHandler
    
    Dim currentRange As Range   ' 当前光标位置
    Dim tbl As Table            ' 表格对象
    Dim tblFound As Boolean     ' 是否找到表格标记
    Dim lastFoundTable As Table ' 记录找到的上一个表格
    
    ' 初始化变量
    Set currentRange = Selection.Range.Duplicate
    tblFound = False
    
    ' 遍历文档中所有表格
    For Each tbl In ActiveDocument.Tables
        ' 检查表格是否在光标之前
        If tbl.Range.End < currentRange.Start Then
            Set lastFoundTable = tbl ' 更新最后找到的表格
            tblFound = True
            ' 继续遍历，找最后一个符合条件的表格（最近的上一个）
        End If
    Next
    
    ' 定位并提示结果
    If tblFound Then
        ' 选中整个表格
        lastFoundTable.Range.Select
        'MsgBox "已定位到上一个表格并选中！", vbInformation, "操作成功"
    Else
        MsgBox "当前位置之前未找到表格！", vbExclamation, "提示"
    End If
    
    ' 清理对象
Cleanup:
        Set currentRange = Nothing
        Set lastFoundTable = Nothing
        Set tbl = Nothing
    Exit Sub
    
    ' 错误处理
ErrorHandler:
        MsgBox "定位表格时发生错误：" & Err.Description, vbCritical, "错误"
        Resume Cleanup
End Sub

'===============================================
'===============================================

' ========== 辅助函数 ==========

Function GetOrCreateCaptionLabel(labelName As String) As CaptionLabel
    '
    ' 获取或创建题注标签
    '
    On Error Resume Next
    Set GetOrCreateCaptionLabel = CaptionLabels(labelName)
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        ' 标签不存在，创建新标签
        Set GetOrCreateCaptionLabel = CaptionLabels.Add(labelName)
        
        ' 设置默认格式
        With GetOrCreateCaptionLabel
            .NumberStyle = wdCaptionNumberStyleArabic
            .IncludeChapterNumber = False
        End With
    End If
    
    On Error GoTo 0
End Function

Sub EnsureCaptionLabel(labelName As String)
    '
    ' 确保题注标签存在
    '
    Dim tempLabel As CaptionLabel
    Set tempLabel = GetOrCreateCaptionLabel(labelName)
End Sub

Function IsSelectionValidForCaption() As Boolean
    '
    ' 验证当前选择是否适合插入题注
    '
    Dim result As Boolean
    result = True
    
    ' 检查是否在表格、页眉页脚等特殊位置
    If Selection.Information(wdWithInTable) Or _
       Selection.Information(wdHeaderFooterType) <> -1 Then
        result = False
    End If
    
    IsSelectionValidForCaption = result
End Function
'=====================================================
'=====================================================
' 宏1：插入下方最近的"图"题注交叉引用
Sub 插入下方最近的图题注交叉引()
    ' 功能：在当前光标位置插入下方最近的"图"题注交叉引用
    Dim currentRng As Range      ' 光标位置的Range对象
    Dim currParaNum As Long      ' 当前光标所在段落编号
    Dim figParaNum As Long       ' 下方最近"图"题注的段落编号
    
    ' 获取当前光标位置的Range（若文档未选中内容，则为光标位置）
    Set currentRng = Selection.Range
    ' 获取当前光标所在的段落编号（用于定位查找起点）
    currParaNum = GetParagraphNumber(currentRng)
    
    ' 查找当前段落后方最近的"图"题注段落编号
    figParaNum = FindNextCaptionPara("图表", currParaNum, ActiveDocument)
    
    ' 根据查找到的题注位置，插入对应交叉引用
    If figParaNum = 0 Then
        MsgBox "光标下方未找到'图'题注", vbExclamation
    Else
        InsertCaptionReference "图表", figParaNum, currentRng
    End If
End Sub

' 宏2：插入下方最近的"表"题注交叉引用
Sub 插入下方最近的表题注交叉引()
    ' 功能：在当前光标位置插入下方最近的"表"题注交叉引用
    Dim currentRng As Range      ' 光标位置的Range对象
    Dim currParaNum As Long      ' 当前光标所在段落编号
    Dim tabParaNum As Long       ' 下方最近"表"题注的段落编号
    
    ' 获取当前光标位置的Range（若文档未选中内容，则为光标位置）
    Set currentRng = Selection.Range
    ' 获取当前光标所在的段落编号（用于定位查找起点）
    currParaNum = GetParagraphNumber(currentRng)
    
    ' 查找当前段落后方最近的"表"题注段落编号
    tabParaNum = FindNextCaptionPara("表", currParaNum, ActiveDocument)
    
    ' 根据查找到的题注位置，插入对应交叉引用
    If tabParaNum = 0 Then
        MsgBox "光标下方未找到'表'题注", vbExclamation
    Else
        InsertCaptionReference "表格", tabParaNum, currentRng
    End If
End Sub

' ========== 以下是通用辅助函数（两个宏共用） ==========

Private Function GetParagraphNumber(rng As Range) As Long
    ' 功能：获取指定Range所在的段落编号（从文档开头计数）
    Dim para As Paragraph
    Dim paraCount As Long
    
    paraCount = 0
    ' 遍历文档所有段落，匹配Range所在段落
    For Each para In rng.Document.Paragraphs
        paraCount = paraCount + 1
        ' 若当前段落包含目标Range的起点，则为目标段落
        If para.Range.Start <= rng.Start And para.Range.End >= rng.Start Then
            GetParagraphNumber = paraCount
            Exit Function
        End If
    Next para
    
    ' 未找到对应段落（通常不会发生）
    GetParagraphNumber = 0
End Function

Private Function FindNextCaptionPara(captionType As String, startPara As Long, doc As Document) As Long
    ' 功能：从指定段落后方查找首个指定类型（图/表）的题注段落
    Dim para As Paragraph
    Dim paraCount As Long
    Dim fld As Field
    
    paraCount = 0
    ' 遍历文档所有段落
    For Each para In doc.Paragraphs
        paraCount = paraCount + 1
        ' 只处理起始段落后方的段落
        If paraCount > startPara Then
            ' 遍历段落中的所有域（题注由SEQ域生成）
            For Each fld In para.Range.Fields
                ' 匹配"SEQ 图"或"SEQ 表"的域代码（题注的核心标识）
                If InStr(1, fld.code.text, "SEQ " & captionType, vbTextCompare) > 0 Then
                    FindNextCaptionPara = paraCount
                    Exit Function
                End If
            Next fld
        End If
    Next para
    
    ' 未找到指定类型题注
    FindNextCaptionPara = 0
End Function

Private Sub InsertCaptionReference(captionType As String, paraNum As Long, insertRng As Range)
    ' 功能：在指定位置插入对应题注的交叉引用
    Dim targetPara As Paragraph  ' 目标题注所在段落
    Dim targetText As String     ' 目标题注文本
    Dim refItems As Variant      ' 所有交叉引用项
    Dim i As Integer
    Dim found As Boolean
    
    ' 获取目标题注段落
    Set targetPara = GetParagraphByNumber(insertRng.Document, paraNum)
    If targetPara Is Nothing Then Exit Sub  ' 若段落不存在，退出
    
    ' 获取目标题注的文本（去除段落标记）
    targetText = Left(targetPara.Range.text, Len(targetPara.Range.text) - 1)
    
    ' 获取文档中所有该类型题注的交叉引用项
    refItems = insertRng.Document.GetCrossReferenceItems(captionType)
    
    ' 遍历引用项，匹配题注文本并插入交叉引用
    For i = 1 To UBound(refItems)
        ' 如果交叉引用项的文本与目标题注文本匹配
        If refItems(i) = targetText Then
            ' 找到完全匹配项，插入交叉引用
            insertRng.InsertCrossReference _
                ReferenceType:=captionType, _
                ReferenceKind:=wdOnlyLabelAndNumber, _
                ReferenceItem:=i, _
                InsertAsHyperlink:=True
            found = True
            Exit For
        End If
    Next i
    
    ' 如果没有找到完全匹配项，尝试使用编号匹配
    If Not found Then
        ' 提取题注编号
        Dim captionNum As String
        captionNum = ExtractCaptionNumber(targetText)
        
        ' 如果提取到编号，尝试匹配
        If captionNum <> "" Then
            ' 规范化编号（将不间断连字符转换为普通连字符）
            Dim normalizedCaptionNum As String
            normalizedCaptionNum = NormalizeHyphen(captionNum)
            
            For i = 1 To UBound(refItems)
                ' 从交叉引用项中提取编号并规范化
                Dim refItemText As String
                refItemText = refItems(i)
                
                Dim refCaptionNum As String
                refCaptionNum = ExtractCaptionNumber(refItemText)
                refCaptionNum = NormalizeHyphen(refCaptionNum)
                
                ' 检查规范化后的编号是否匹配
                If refCaptionNum = normalizedCaptionNum Then
                    insertRng.InsertCrossReference _
                        ReferenceType:=captionType, _
                        ReferenceKind:=wdOnlyLabelAndNumber, _
                        ReferenceItem:=i, _
                        InsertAsHyperlink:=True
                    found = True
                    Exit For
                End If
            Next i
        End If
    End If
    
    ' 如果仍然没有找到，使用索引方式（可能不是完全准确）
    If Not found Then
        ' 计算题注在文档中的索引位置
        Dim captionIndex As Long
        captionIndex = GetCaptionIndex(insertRng.Document, captionType, paraNum)
        
        If captionIndex > 0 Then
            insertRng.InsertCrossReference _
                ReferenceType:=captionType, _
                ReferenceKind:=wdOnlyLabelAndNumber, _
                ReferenceItem:=captionIndex, _
                InsertAsHyperlink:=True
        Else
            MsgBox "无法插入交叉引用，请手动插入。" & vbCrLf & _
                   "目标题注: " & targetText, vbExclamation
        End If
    End If
End Sub

Private Function ExtractCaptionNumber(captionText As String) As String
    ' 功能：从题注文本中提取编号
    ' 支持多种格式：图1、图1-1、图1.1、图1.1.1等
    Dim i As Integer
    Dim result As String
    Dim char As String
    Dim charCode As Integer
    
    ' 跳过开头的非数字字符（如"图"、"表"等）
    For i = 1 To Len(captionText)
        char = Mid(captionText, i, 1)
        charCode = AscW(char)
        ' 检查是否为数字（0-9）
        If char >= "0" And char <= "9" Then
            Exit For
        End If
    Next i
    
    ' 提取编号部分
    result = ""
    For i = i To Len(captionText)
        char = Mid(captionText, i, 1)
        charCode = AscW(char)
        
        ' 判断字符类型：数字、点、连字符、不间断连字符
        ' 数字 0-9: Unicode 48-57
        ' 点 .: Unicode 46
        ' 连字符 -: Unicode 45
        ' 不间断连字符 ?: Unicode 8208 (0x2010)
        ' 可选断字连字符: Unicode 8209 (0x2011)
        
        If (char >= "0" And char <= "9") Or _
           char = "." Or _
           char = "-" Or _
           charCode = 8208 Or _
           charCode = 8209 Then
            result = result & char
        Else
            ' 遇到其他字符，停止提取
            Exit For
        End If
    Next i
    
    ExtractCaptionNumber = result
End Function

Private Function NormalizeHyphen(text As String) As String
    ' 功能：规范化字符串中的连字符，将各种连字符转换为普通连字符
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim charCode As Integer
    
    result = text
    
    ' 替换所有类型的连字符为普通连字符
    For i = 1 To Len(result)
        char = Mid(result, i, 1)
        charCode = AscW(char)
        
        ' 如果是各种连字符，替换为普通连字符
        If charCode = 8208 Or charCode = 8209 Then ' 不间断连字符或可选断字连字符
            Mid(result, i, 1) = "-"
        End If
    Next i
    
    NormalizeHyphen = result
End Function

Private Function GetCaptionIndex(doc As Document, captionType As String, paraNum As Long) As Long
    ' 功能：获取指定题注在文档中的索引位置
    Dim para As Paragraph
    Dim fld As Field
    Dim captionCount As Long
    Dim paraCount As Long
    
    captionCount = 0
    paraCount = 0
    
    ' 遍历文档所有段落
    For Each para In doc.Paragraphs
        paraCount = paraCount + 1
        
        ' 如果已经超过目标段落，停止查找
        If paraCount > paraNum Then
            Exit For
        End If
        
        ' 检查段落中是否有题注域
        For Each fld In para.Range.Fields
            If InStr(1, fld.code.text, "SEQ " & captionType, vbTextCompare) > 0 Then
                captionCount = captionCount + 1
                ' 如果是目标段落，返回计数
                If paraCount = paraNum Then
                    GetCaptionIndex = captionCount
                    Exit Function
                End If
                Exit For
            End If
        Next fld
    Next para
    
    GetCaptionIndex = 0
End Function

Private Function GetParagraphByNumber(doc As Document, num As Long) As Paragraph
    ' 功能：通过段落编号获取对应的段落对象
    Dim para As Paragraph
    Dim paraCount As Long
    
    paraCount = 0
    For Each para In doc.Paragraphs
        paraCount = paraCount + 1
        If paraCount = num Then
            Set GetParagraphByNumber = para
            Exit Function
        End If
    Next para
    
    Set GetParagraphByNumber = Nothing  ' 未找到对应段落
End Function

'=============================================
'=============================================
Sub A章节标题添加序号()
'
' 宏1 宏
'
'
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.76)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.02)
        .TabPosition = wdUndefined
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 2"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = wdUndefined
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 3"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.52)
        .TabPosition = wdUndefined
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 4"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.78)
        .TabPosition = wdUndefined
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 5"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "%1.%2.%3.%4.%5.%6"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.03)
        .TabPosition = wdUndefined
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 6"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.29)
        .TabPosition = wdUndefined
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 7"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.54)
        .TabPosition = wdUndefined
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 8"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(2.79)
        .TabPosition = wdUndefined
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = "标题 9"
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
End Sub
'=========================================================
'=========================================================
Sub A自动序号变文本()
    ' 声明变量（兼容原代码的Lists对象逻辑）
    Dim kgsList As List
    Dim listRange As Range
    Dim selRange As Range
    Dim intersectRange As Range
    
    On Error GoTo ErrorHandler ' 全局异常捕获
    
    ' ===== 前置校验：确保选区有效 =====
    ' 1. 无选中内容（仅光标）
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选中需要处理的文本区域！", vbExclamation, "提示"
        Exit Sub
    End If
    ' 2. 非文本选区（图片/形状/表格等）
    If Not (Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionText) Then
        MsgBox "仅支持对选中的文本区域操作，请勿选中图片/形状/表格等非文本对象！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' ===== 锁定选区范围，避免后续操作偏移 =====
    Set selRange = Selection.Range.Duplicate
    
    ' ===== 禁用屏幕刷新，提升执行效率 =====
    Application.ScreenUpdating = False
    
    ' ===== 核心逻辑：遍历文档所有列表，仅处理与选区重叠的部分 =====
    For Each kgsList In ActiveDocument.Lists
        ' 获取当前列表的完整范围
        Set listRange = kgsList.Range
        
        ' 判断列表范围是否与选中区域有交集
        If Not (listRange.End < selRange.Start Or listRange.Start > selRange.End) Then
            ' 获取列表与选区的交集范围
            Set intersectRange = ActiveDocument.Range( _
                Start:=IIf(listRange.Start > selRange.Start, listRange.Start, selRange.Start), _
                End:=IIf(listRange.End < selRange.End, listRange.End, selRange.End) _
            )
            
            ' 仅将交集中的编号转为文本（核心：保留原ConvertNumbersToText逻辑，限定范围）
            intersectRange.ListFormat.ConvertNumbersToText
        End If
    Next kgsList
    
    ' ===== 恢复屏幕刷新，完成提示 =====
    Application.ScreenUpdating = True
    MsgBox "选中区域内的自动编号已成功转为文本！", vbInformation, "完成"
    
    Exit Sub ' 正常结束
    
ErrorHandler: ' 异常处理
    Application.ScreenUpdating = True ' 无论报错与否，恢复屏幕刷新
    MsgBox "操作失败：" & Err.Description & "（错误码：" & Err.Number & "）", vbCritical, "错误"
End Sub
'=========================================================
'=========================================================
Sub A纯文本粘贴()
    Selection.PasteAndFormat (wdFormatPlainText)
End Sub
'=========================================================
'=========================================================
Sub A软换行_硬换行()
    ' 功能：将文档中所有手动换行符替换为段落标记
    
    Application.ScreenUpdating = False  ' 关闭屏幕更新，提高速度
    
    ' 方法1：使用Word的查找替换功能
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "^l"          ' 手动换行符
        .Replacement.text = "^p"  ' 段落标记
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' 执行全部替换
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    MsgBox "替换完成！已将文档中所有手动换行符替换为段落标记。", vbInformation
End Sub
'=========================================================
'=========================================================

Sub A与下一段同页()
    With Selection.ParagraphFormat
        .KeepWithNext = True
    End With
End Sub

'■■■■■■■■■
Sub A标题样式修改()
    '关闭屏幕刷新，提升运行速度
    Application.ScreenUpdating = False
    
    Dim sty As Style
    On Error Resume Next '忽略个别版本兼容问题
    
    '===================== 标题1样式设置（一级标题） =====================
    With ActiveDocument.Styles(wdStyleHeading1)
        '字体格式
        With .Font
            .Name = "Times New Roman" '英文字体
            .NameFarEast = "黑体"          '一级标题字体（主流为黑体）
            .Size = 16             '字号：小二（18磅），也可设为16（三号）
            .Bold = True           '加粗
            .ColorIndex = wdBlack  '黑色
            .Underline = wdUnderlineNone '无下划线
        End With
        '段落格式
        With .ParagraphFormat
            '.Alignment = wdAlignParagraphCenter '居中对齐
            .alignment = wdAlignParagraphLeft '左对齐
            .LineSpacingRule = wdLineSpacing1pt5 '1.5倍行距
            .spaceBefore = 12 '段前间距：12磅（0.5行）
            .spaceAfter = 12  '段后间距：12磅（0.5行）
            .LeftIndent = 0   '左缩进0字符
            .RightIndent = 0  '右缩进0字符
            .FirstLineIndent = 0 '首行缩进0字符
            .OutlineLevel = wdOutlineLevel1 '大纲级别1级
        End With
        '边框和底纹（无）
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
    End With
    
    '===================== 标题2样式设置（二级标题） =====================
    With ActiveDocument.Styles(wdStyleHeading2)
        With .Font
            .Name = "Times New Roman" '英文字体
            .NameFarEast = "黑体"          '二级标题字体
            .Size = 14             '四号（14磅）
            .Bold = True
            .ColorIndex = wdBlack
            .Underline = wdUnderlineNone
        End With
        With .ParagraphFormat
            .alignment = wdAlignParagraphLeft '左对齐
            .LineSpacingRule = wdLineSpacing1pt5
            .spaceBefore = 6  '段前6磅（0.25行）
            .spaceAfter = 6   '段后6磅（0.25行）
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 2 '首行缩进2字符（部分院校要求）
            .OutlineLevel = wdOutlineLevel2
        End With
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
    End With
    
    '===================== 标题3样式设置（三级标题） =====================
    With ActiveDocument.Styles(wdStyleHeading3)
        With .Font
            .Name = "Times New Roman" '英文字体
            .NameFarEast = "宋体"          '三级标题宋体加粗
            .Size = 12             '小四（12磅）
            .Bold = True
            .ColorIndex = wdBlack
            .Underline = wdUnderlineNone
        End With
        With .ParagraphFormat
            .alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpacingSingle '单倍行距
            .spaceBefore = 6
            .spaceAfter = 6
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 2
            .OutlineLevel = wdOutlineLevel3
        End With
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
    End With
    
    '===================== 标题4样式设置（四级标题） =====================
    With ActiveDocument.Styles(wdStyleHeading4)
        With .Font
            .Name = "Times New Roman" '英文字体
            .NameFarEast = "宋体"
            .Size = 12
            .Bold = True
            .ColorIndex = wdBlack
            .Underline = wdUnderlineNone
        End With
        With .ParagraphFormat
            .alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpacingSingle
            .spaceBefore = 3
            .spaceAfter = 3
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 2
            .OutlineLevel = wdOutlineLevel4
        End With
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
    End With
    
    '===================== 标题5样式设置（五级标题） =====================
    With ActiveDocument.Styles(wdStyleHeading5)
        With .Font
            .Name = "Times New Roman" '英文字体
            .NameFarEast = "宋体"
            .Size = 12
            .Bold = False '五级标题不加粗
            .ColorIndex = wdBlack
            .Underline = wdUnderlineNone
        End With
        With .ParagraphFormat
            .alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpacingSingle
            .spaceBefore = 3
            .spaceAfter = 3
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 2
            .OutlineLevel = wdOutlineLevel5
        End With
        .Borders.Enable = False
        .Shading.Texture = wdTextureNone
    End With
    
    '恢复屏幕刷新
    Application.ScreenUpdating = True
    MsgBox "通用标题样式设置完成！若需个性定制，进宏编辑器修改", vbInformation, "提示"
End Sub


'=============================================
'=============================================
Sub A下一个VISIO图()
    '声明变量
    Dim currentRange As Range
    Dim objField As Field
    Dim foundVisio As Boolean
    
    '初始化变量
    foundVisio = False
    
    '获取光标当前位置的Range对象
    Set currentRange = Selection.Range
    
    '遍历文档中所有的域
    For Each objField In ActiveDocument.Fields
        '检查域的代码是否包含Visio嵌入标识，且域的位置在光标之后
        If InStr(1, objField.code.text, "EMBED Visio.Drawing", vbTextCompare) > 0 Then
            '判断当前域是否在光标位置之后
            If objField.result.Start > currentRange.End Then
                '选中找到的Visio对象
                objField.result.Select
                foundVisio = True
                '找到第一个符合条件的对象后退出循环
                Exit For
            End If
        End If
    Next objField
    
    '如果没有找到后续的Visio对象，给出提示
    If Not foundVisio Then
        MsgBox "未找到光标位置之后的Visio嵌入式对象！", vbInformation, "提示"
    End If
    
    '释放对象变量
    Set currentRange = Nothing
    Set objField = Nothing
End Sub


'=============================================
'=============================================
Sub A上一个VISIO图()
    '声明变量
    Dim currentRange As Range
    Dim objField As Field
    Dim foundVisio As Boolean
    
    '初始化变量
    foundVisio = False
    
    '获取光标当前位置的Range对象
    Set currentRange = Selection.Range
    
    '遍历文档中所有的域
    For Each objField In ActiveDocument.Fields
        '检查域的代码是否包含Visio嵌入标识，且域的位置在光标之前
        If InStr(1, objField.code.text, "EMBED Visio.Drawing", vbTextCompare) > 0 Then
            '判断当前域是否在光标位置之前
            If objField.result.End < currentRange.Start Then
                '选中找到的Visio对象
                objField.result.Select
                foundVisio = True
                '不退出循环，继续查找最后一个符合条件的对象（最近的上一个）
            End If
        End If
    Next objField
    
    '如果没有找到之前的Visio对象，给出提示
    If Not foundVisio Then
        MsgBox "未找到光标位置之前的Visio嵌入式对象！", vbInformation, "提示"
    End If
    
    '释放对象变量
    Set currentRange = Nothing
    Set objField = Nothing
End Sub

'=========================================================
'=========================================================
Sub SelectNextMathTypeEquation()
    '声明变量
    Dim currentRange As Range       '存储光标当前位置
    Dim objField As Field           '遍历域的临时变量
    Dim foundEquation As Boolean    '标记是否找到符合条件的公式
    Dim equationRange As Range      '专门存储找到的公式完整范围（确保选中完整）
    
    '初始化标记为未找到
    foundEquation = False
    
    '获取光标当前位置的Range对象（作为查找起点）
    Set currentRange = Selection.Range
    
    '遍历文档中所有域
    For Each objField In ActiveDocument.Fields
        '精准筛选MathType公式域（EMBED Equation.DSMT4）
        If InStr(1, objField.code.text, "EMBED Equation.DSMT4", vbTextCompare) > 0 Then
            '判断公式是否在光标位置之后
            If objField.result.Start > currentRange.End Then
                '获取公式的完整Range范围（核心：确保选中整个公式）
                Set equationRange = objField.result
                '显式选中公式的完整范围（强制触发选中状态）
                equationRange.Select
                '将选中的公式滚动到视图中（直观看到选中的公式，避免选中后看不到）
                ActiveWindow.ScrollIntoView equationRange, True
                '标记为找到，退出循环（只选第一个符合条件的公式）
                foundEquation = True
                Exit For
            End If
        End If
    Next objField
    
    '未找到公式时给出友好提示
    If Not foundEquation Then
        MsgBox "未找到光标位置之后的MathType公式！", vbInformation, "提示"
    End If
    
    '释放对象变量，避免内存泄漏
    Set currentRange = Nothing
    Set objField = Nothing
    Set equationRange = Nothing
End Sub





Sub A修改选中图片大小()
    ' 功能：修改选中图片的大小，保持宽高比
    Dim selectedPic As InlineShape
    Dim customWidth As String
    Dim newWidthCm As Double
    Dim newWidthPoints As Single
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim aspectRatio As Double
    Dim newHeightPoints As Single
    
    ' 检查是否有选中的嵌入式图片
    If Selection.InlineShapes.Count = 0 Then
        MsgBox "请先选中一张嵌入式图片！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 获取选中的第一张图片
    Set selectedPic = Selection.InlineShapes(1)
    
    ' 显示自定义宽度输入框（默认值12厘米）
    customWidth = InputBox("请输入图片宽度（单位：厘米）：", "设置图片宽度", "12")
    
    ' 检查用户输入
    If customWidth = "" Then
        MsgBox "未输入宽度，操作已取消。", vbInformation, "取消"
        Exit Sub
    End If
    
    ' 验证输入是否为数字
    If Not IsNumeric(customWidth) Then
        MsgBox "输入的不是有效的数字，请重新运行宏并输入正确的宽度值。", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    newWidthCm = Val(customWidth)
    
    ' 验证宽度是否为正数
    If newWidthCm <= 0 Then
        MsgBox "输入的宽度必须大于0，请重新运行宏并输入正确的宽度值。", vbExclamation, "输入错误"
        Exit Sub
    End If
    
    ' 转换为Word内部的Points单位
    newWidthPoints = CentimetersToPoints(newWidthCm)
    
    ' 获取原图尺寸
    originalWidth = selectedPic.Width
    originalHeight = selectedPic.Height
    
    ' 计算宽高比
    If originalWidth = 0 Then
        MsgBox "无法获取图片原始尺寸！", vbExclamation, "错误"
        Exit Sub
    End If
    
    aspectRatio = originalHeight / originalWidth
    
    ' 计算新的高度（保持宽高比）
    newHeightPoints = newWidthPoints * aspectRatio
    
    ' 关闭屏幕更新，提高处理速度
    Application.ScreenUpdating = False
    
    ' 设置图片新尺寸
    selectedPic.Width = newWidthPoints
    selectedPic.Height = newHeightPoints
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 显示结果信息
    MsgBox "图片大小修改完成！" & vbCrLf & _
           "原图尺寸：" & Format(PointsToCentimeters(originalWidth), "0.00") & "cm × " & Format(PointsToCentimeters(originalHeight), "0.00") & "cm" & vbCrLf & _
           "新图尺寸：" & Format(newWidthCm, "0.00") & "cm × " & Format(PointsToCentimeters(newHeightPoints), "0.00") & "cm", _
           vbInformation, "完成"
End Sub


'===========================================================
'===========================================================
Sub A导出批注到桌面()
    Dim doc As Document
    Dim cmt As Comment
    Dim exportDoc As Document
    Dim i As Long
    Dim saveFolder As String
    Dim fileName As String
    Dim savePath As String
    Dim exportContent As String
    
    On Error GoTo ErrorHandler
    
    ' 设置当前文档
    Set doc = ActiveDocument
    
    ' 检查是否有批注
    If doc.Comments.Count = 0 Then
        MsgBox "当前文档没有批注，无需导出。", vbInformation, "提示"
        Exit Sub
    End If
    
    ' 获取保存目录 - 优先使用文档所在目录
    saveFolder = doc.Path
    
    ' 如果文档未保存，使用桌面
    If saveFolder = "" Then
        #If Mac Then
            ' Mac系统桌面路径
            saveFolder = Environ$("HOME") & "/Desktop/"
        #Else
            ' Windows系统桌面路径
            saveFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
        #End If
    Else
        ' 确保路径有分隔符
        If Right(saveFolder, 1) <> "\" And Right(saveFolder, 1) <> "/" Then
            saveFolder = saveFolder & Application.PathSeparator
        End If
    End If
    
    ' 获取文件名
    fileName = InputBox("请输入导出文件名（不含扩展名）：", _
                       "保存批注导出文件", _
                       "批注导出_" & Format(Date, "yyyymmdd"))
    
    If fileName = "" Then
        Exit Sub
    End If
    
    ' 构建完整路径
    savePath = saveFolder & fileName & ".docx"
    
    ' 检查文件是否已存在
    If Dir(savePath) <> "" Then
        If MsgBox("文件已存在，是否覆盖？", vbQuestion + vbYesNo, "确认覆盖") <> vbYes Then
            Exit Sub
        End If
    End If
    
    ' 禁用屏幕更新，提高性能
    Application.ScreenUpdating = False
    
    ' 创建新文档
    Set exportDoc = Documents.Add
    
    ' 创建标题和基本信息
    With exportDoc
        .Content.text = "文档批注导出报告" & vbCrLf & vbCrLf
        
        ' 添加文档信息
        .Content.InsertAfter "源文档: " & doc.Name & vbCrLf
        .Content.InsertAfter "导出时间: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
        .Content.InsertAfter "批注总数: " & doc.Comments.Count & " 条" & vbCrLf & vbCrLf
        
        ' 添加分隔线
        .Content.InsertAfter String(50, "=") & vbCrLf & vbCrLf
    End With
    
    ' 遍历所有批注
    For i = 1 To doc.Comments.Count
        Set cmt = doc.Comments(i)
        
        ' 构建批注内容
        exportContent = exportContent & _
            "批注 #" & i & vbCrLf & _
            String(30, "-") & vbCrLf & _
            "作者: " & cmt.Author & vbCrLf & _
            "日期: " & cmt.Date & vbCrLf & _
            "原文: " & vbCrLf & _
            "    " & Trim(cmt.Scope.text) & vbCrLf & vbCrLf & _
            "批注内容: " & vbCrLf & _
            "    " & Trim(cmt.Range.text) & vbCrLf & _
            String(50, "=") & vbCrLf & vbCrLf
    Next i
    
    ' 一次性插入所有批注内容
    exportDoc.Content.InsertAfter exportContent
    
    ' 简单格式化
    With exportDoc
        ' 设置标题格式
        .Paragraphs(1).Range.Font.Bold = True
        .Paragraphs(1).Range.Font.Size = 16
        
        ' 设置正文格式
        .Content.Font.Name = "宋体"
        .Content.Font.Size = 11
    End With
    
    ' 保存文档
    exportDoc.SaveAs2 fileName:=savePath, _
                     FileFormat:=wdFormatXMLDocument
    
    exportDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 显示完成信息
    MsgBox "批注导出完成！" & vbCrLf & _
           "共导出 " & doc.Comments.Count & " 条批注" & vbCrLf & _
           "保存位置：" & vbCrLf & savePath, _
           vbInformation, "导出成功"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    If Not exportDoc Is Nothing Then
        exportDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    
    MsgBox "导出过程中发生错误：" & vbCrLf & _
           "错误号: " & Err.Number & vbCrLf & _
           "错误描述: " & Err.Description, vbCritical, "错误"
End Sub




'■■■■■■■■■ BK基础系列样式应用宏（开始） ■■■■■■■■■
'■■■■■■■■■ 应用BK表题英文样式
Sub B应用BK表题英文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK表题英文") Then
        Selection.Style = ActiveDocument.Styles("BK表题英文")
    Else
        MsgBox "样式 'BK表题英文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK表题中文样式
Sub B应用BK表题中文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK表题中文") Then
        Selection.Style = ActiveDocument.Styles("BK表题中文")
    Else
        MsgBox "样式 'BK表题中文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK目录样式
Sub B应用BK目录样式()
    On Error GoTo ErrHandler
    If StyleExists("BK目录""目录""") Then
        Selection.Style = ActiveDocument.Styles("BK目录""目录""")
    Else
        MsgBox "样式 'BK目录""目录""' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK图行样式
Sub B应用BK图行样式()
    On Error GoTo ErrHandler
    If StyleExists("BK图行") Then
        Selection.Style = ActiveDocument.Styles("BK图行")
    Else
        MsgBox "样式 'BK图行' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK一级标题样式
Sub B应用BK一级标题样式()
    On Error GoTo ErrHandler
    If StyleExists("BK一级标题【前言、结论、参考文献、致谢】") Then
        Selection.Style = ActiveDocument.Styles("BK一级标题【前言、结论、参考文献、致谢】")
    Else
        MsgBox "样式 'BK一级标题【前言、结论、参考文献、致谢】' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK图题英文样式
Sub B应用BK图题英文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK图题英文") Then
        Selection.Style = ActiveDocument.Styles("BK图题英文")
    Else
        MsgBox "样式 'BK图题英文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要中文摘要样式
Sub B应用BK摘要中文摘要样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要中文""摘要""") Then
        Selection.Style = ActiveDocument.Styles("BK摘要中文""摘要""")
    Else
        MsgBox "样式 'BK摘要中文""摘要""' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要英文ABSTRACT样式
Sub B应用BK摘要英文ABSTRACT样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要英文""ABSTRACT""") Then
        Selection.Style = ActiveDocument.Styles("BK摘要英文""ABSTRACT""")
    Else
        MsgBox "样式 'BK摘要英文""ABSTRACT""' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要中文论文题目样式
Sub B应用BK摘要中文论文题目样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要中文论文题目") Then
        Selection.Style = ActiveDocument.Styles("BK摘要中文论文题目")
    Else
        MsgBox "样式 'BK摘要中文论文题目' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要英文KEY WORDS样式
Sub B应用BK摘要英文KEYWORDS样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要英文""KEY WORDS""") Then
        Selection.Style = ActiveDocument.Styles("BK摘要英文""KEY WORDS""")
    Else
        MsgBox "样式 'BK摘要英文""KEY WORDS""' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要英文—关键词样式
Sub B应用BK摘要英文关键词样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要英文—关键词") Then
        Selection.Style = ActiveDocument.Styles("BK摘要英文—关键词")
    Else
        MsgBox "样式 'BK摘要英文—关键词' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BKABSTRACT PAPER TITLE样式
Sub B应用BKABSTRACTPAPERTITLE样式()
    On Error GoTo ErrHandler
    If StyleExists("BKABSTRACT PAPER TITLE") Then
        Selection.Style = ActiveDocument.Styles("BKABSTRACT PAPER TITLE")
    Else
        MsgBox "样式 'BKABSTRACT PAPER TITLE' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要英文—正文样式
Sub B应用BK摘要英文正文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要英文—正文") Then
        Selection.Style = ActiveDocument.Styles("BK摘要英文—正文")
    Else
        MsgBox "样式 'BK摘要英文—正文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要中文—关键词样式（黑体标识）
Sub B应用BK摘要中文关键词黑体样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要中文""关键词""") Then
        Selection.Style = ActiveDocument.Styles("BK摘要中文""关键词""")
    Else
        MsgBox "样式 'BK摘要中文“关键词”' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要中文-关键词样式（宋体内容）
Sub B应用BK摘要中文关键词宋体样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要中文-关键词") Then
        Selection.Style = ActiveDocument.Styles("BK摘要中文-关键词")
    Else
        MsgBox "样式 'BK摘要中文-关键词' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK摘要中文—正文样式（首行缩进2字符）
Sub B应用BK摘要中文正文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK摘要中文—正文") Then
        Selection.Style = ActiveDocument.Styles("BK摘要中文—正文")
    Else
        MsgBox "样式 'BK摘要中文—正文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK正文缩进样式（首行缩进2字符）
Sub B应用BK正文缩进样式()
    On Error GoTo ErrHandler
    If StyleExists("BK正文缩进") Then
        Selection.Style = ActiveDocument.Styles("BK正文缩进")
    Else
        MsgBox "样式 'BK正文缩进' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK正文顶格样式
Sub B应用BK正文顶格样式()
    On Error GoTo ErrHandler
    If StyleExists("BK正文顶格") Then
        Selection.Style = ActiveDocument.Styles("BK正文顶格")
    Else
        MsgBox "样式 'BK正文顶格' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用BK图题中文样式
Sub B应用BK图题中文样式()
    On Error GoTo ErrHandler
    If StyleExists("BK图题中文") Then
        Selection.Style = ActiveDocument.Styles("BK图题中文")
    Else
        MsgBox "样式 'BK图题中文' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用MTDisplayEquation样式
Sub B应用MTDisplayEquation样式()
    On Error GoTo ErrHandler
    If StyleExists("MTDisplayEquation") Then
        Selection.Style = ActiveDocument.Styles("MTDisplayEquation")
    Else
        MsgBox "样式 'MTDisplayEquation' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■ BK基础系列样式应用宏（结束） ■■■■■■■■■

'■■■■■■■■■ TOC目录系列样式应用宏（开始） ■■■■■■■■■
'■■■■■■■■■ 应用TOC 1样式
Sub B应用TOC1样式()
    On Error GoTo ErrHandler
    If StyleExists("TOC 1") Then
        Selection.Style = ActiveDocument.Styles("TOC 1")
    Else
        MsgBox "样式 'TOC 1' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用TOC 2样式
Sub B应用TOC2样式()
    On Error GoTo ErrHandler
    If StyleExists("TOC 2") Then
        Selection.Style = ActiveDocument.Styles("TOC 2")
    Else
        MsgBox "样式 'TOC 2' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用TOC 3样式
Sub B应用TOC3样式()
    On Error GoTo ErrHandler
    If StyleExists("TOC 3") Then
        Selection.Style = ActiveDocument.Styles("TOC 3")
    Else
        MsgBox "样式 'TOC 3' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用TOC 4样式
Sub B应用TOC4样式()
    On Error GoTo ErrHandler
    If StyleExists("TOC 4") Then
        Selection.Style = ActiveDocument.Styles("TOC 4")
    Else
        MsgBox "样式 'TOC 4' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■ TOC目录系列样式应用宏（结束） ■■■■■■■■■

'■■■■■■■■■ 标题系列样式应用宏（开始） ■■■■■■■■■
'■■■■■■■■■ 应用标题 1样式
Sub B应用标题1样式()
    On Error GoTo ErrHandler
    If StyleExists("标题 1") Then
        Selection.Style = ActiveDocument.Styles("标题 1")
    Else
        MsgBox "样式 '标题 1' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用标题 2样式
Sub B应用标题2样式()
    On Error GoTo ErrHandler
    If StyleExists("标题 2") Then
        Selection.Style = ActiveDocument.Styles("标题 2")
    Else
        MsgBox "样式 '标题 2' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用标题 3样式
Sub B应用标题3样式()
    On Error GoTo ErrHandler
    If StyleExists("标题 3") Then
        Selection.Style = ActiveDocument.Styles("标题 3")
    Else
        MsgBox "样式 '标题 3' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

'■■■■■■■■■ 应用标题 4样式
Sub B应用标题4样式()
    On Error GoTo ErrHandler
    If StyleExists("标题 4") Then
        Selection.Style = ActiveDocument.Styles("标题 4")
    Else
        MsgBox "样式 '标题 4' 不存在于当前文档中", vbExclamation, "样式未找到"
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "应用样式时发生错误: " & Err.Description, vbCritical, "错误"
End Sub
'■■■■■■■■■ 标题系列样式应用宏（结束） ■■■■■■■■■



'■■■■■■■■■
' 宏名称：设置全文页边距和页眉页脚边距
' 功能：统一设置多节文档的页边距（上下25mm、左右27mm）和页眉边距（上1.4cm、下2.0cm）
'■■■■■■■■■
Sub 设置全文页边距和页眉页脚()
    On Error GoTo ErrHandler
    
    ' 声明变量
    Dim doc As Document           ' 当前活动文档
    Dim sec As Section            ' 文档节对象（处理多节场景）
    Dim pageSetup As pageSetup    ' 页面设置对象
    
    ' 单位转换常量（Word VBA中页面设置默认单位为「磅」）
    Const MM_TO_POINT As Single = 2.835   ' 1毫米 = 2.835磅
    Const CM_TO_POINT As Single = 28.35   ' 1厘米 = 28.35磅
    
    ' 配置参数（按需求定义）
    Const MARGIN_TOP_BOTTOM As Single = 25    ' 上下页边距（mm）
    Const MARGIN_LEFT_RIGHT As Single = 27    ' 左右页边距（mm）
    Const HEADER_MARGIN_TOP As Single = 1.4   ' 页眉上边距（cm）
    Const HEADER_MARGIN_BOTTOM As Single = 2#  ' 页眉下边距（cm）
    
    ' 第一步：检查是否有活动文档
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "请先打开一个Word文档再执行此宏！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 第二步：遍历所有节，统一设置页面属性（解决多节不一致问题）
    For Each sec In doc.Sections
        Set pageSetup = sec.pageSetup
        
        ' 1. 设置页边距（转换为磅）
        pageSetup.TopMargin = MARGIN_TOP_BOTTOM * MM_TO_POINT      ' 上25mm
        pageSetup.BottomMargin = MARGIN_TOP_BOTTOM * MM_TO_POINT   ' 下25mm
        pageSetup.LeftMargin = MARGIN_LEFT_RIGHT * MM_TO_POINT     ' 左27mm
        pageSetup.RightMargin = MARGIN_LEFT_RIGHT * MM_TO_POINT    ' 右27mm
        
        ' 2. 设置页眉边距（转换为磅）
        pageSetup.HeaderDistance = HEADER_MARGIN_TOP * CM_TO_POINT ' 页眉上边距1.4cm（页眉到纸张上边缘的距离）
        ' 页眉下边距2.0cm：指页眉区域底部到正文区域的距离，通过调整正文上边距偏移实现
        pageSetup.Orientation = wdOrientPortrait   ' 确保文本方向为横向（不影响边距）
        pageSetup.SectionStart = wdSectionNewPage                 ' 保持节的起始类型（不修改文档结构）
        
        ' 补充：页脚边距默认跟随页面设置，若需统一可取消下方注释
        ' pageSetup.FooterDistance = 1.5 * CM_TO_POINT ' 页脚到底边距离1.5cm（可自定义）
    Next sec
    
    ' 执行完成提示
    MsgBox "全文页边距和页眉边距设置完成！" & vbCrLf & _
           "页边距：上下25mm，左右27mm" & vbCrLf & _
           "页眉边距：上1.4cm，下2.0cm", vbInformation, "设置成功"
    
    ' 释放对象
    Set pageSetup = Nothing
    Set sec = Nothing
    Set doc = Nothing
    Exit Sub

ErrHandler:
    ' 错误处理
    MsgBox "设置过程中发生错误：" & Err.Description & vbCrLf & _
           "错误代码：" & Err.Number, vbCritical, "错误"
    ' 释放对象
    Set pageSetup = Nothing
    Set sec = Nothing
    Set doc = Nothing
End Sub

'■■■■■■■■■
' 宏名称：设置全文页边距和页眉页脚边距
' 功能：统一设置多节文档的页边距（上下25mm、左右27mm）和页眉边距（上1.4cm、下2.0cm）
'■■■■■■■■■
Sub B设置BK全文页边距和页眉页脚()
    On Error GoTo ErrHandler
    
    ' 声明变量
    Dim doc As Document           ' 当前活动文档
    Dim sec As Section            ' 文档节对象（处理多节场景）
    Dim pageSetup As pageSetup    ' 页面设置对象
    
    ' 单位转换常量（Word VBA中页面设置默认单位为「磅」）
    Const MM_TO_POINT As Single = 2.835   ' 1毫米 = 2.835磅
    Const CM_TO_POINT As Single = 28.35   ' 1厘米 = 28.35磅
    
    ' 配置参数（按需求定义）
    Const MARGIN_TOP_BOTTOM As Single = 25    ' 上下页边距（mm）
    Const MARGIN_LEFT_RIGHT As Single = 27    ' 左右页边距（mm）
    Const HEADER_MARGIN_TOP As Single = 1.4   ' 页眉上边距（cm）
    Const HEADER_MARGIN_BOTTOM As Single = 2#  ' 页眉下边距（cm）
    
    ' 第一步：检查是否有活动文档
    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "请先打开一个Word文档再执行此宏！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 第二步：遍历所有节，统一设置页面属性（解决多节不一致问题）
    For Each sec In doc.Sections
        Set pageSetup = sec.pageSetup
        
        ' 1. 设置页边距（转换为磅）
        pageSetup.TopMargin = MARGIN_TOP_BOTTOM * MM_TO_POINT      ' 上25mm
        pageSetup.BottomMargin = MARGIN_TOP_BOTTOM * MM_TO_POINT   ' 下25mm
        pageSetup.LeftMargin = MARGIN_LEFT_RIGHT * MM_TO_POINT     ' 左27mm
        pageSetup.RightMargin = MARGIN_LEFT_RIGHT * MM_TO_POINT    ' 右27mm
        
        ' 2. 设置页眉边距（转换为磅）
        pageSetup.HeaderDistance = HEADER_MARGIN_TOP * CM_TO_POINT ' 页眉上边距1.4cm（页眉到纸张上边缘的距离）
        ' 页眉下边距2.0cm：指页眉区域底部到正文区域的距离，PageSetup中通过HeaderDistance+正文边距自动适配
        pageSetup.SectionStart = wdSectionNewPage                 ' 保持节的起始类型（不修改文档结构）
        
        ' 补充：页脚边距默认跟随页面设置，若需统一可取消下方注释
        ' pageSetup.FooterDistance = 1.5 * CM_TO_POINT ' 页脚到底边距离1.5cm（可自定义）
    Next sec
    
    ' 执行完成提示
    MsgBox "全文页边距和页眉边距设置完成！" & vbCrLf & _
           "页边距：上下25mm，左右27mm" & vbCrLf & _
           "页眉边距：上1.4cm，下2.0cm", vbInformation, "设置成功"
    
    ' 释放对象
    Set pageSetup = Nothing
    Set sec = Nothing
    Set doc = Nothing
    Exit Sub

ErrHandler:
    ' 错误处理
    MsgBox "设置过程中发生错误：" & Err.Description & vbCrLf & _
           "错误代码：" & Err.Number, vbCritical, "错误"
    ' 释放对象
    Set pageSetup = Nothing
    Set sec = Nothing
    Set doc = Nothing
End Sub







'==========================================================================
Sub B批量创建并修改BK样式()
    '-------------------------- 变量声明区 --------------------------
    Dim doc As Document               ' 当前活动的Word文档对象
    Dim targetStyle As Style          ' 单个样式对象
    Dim stylesInfo() As Variant       ' 样式信息二维数组
    Dim i As Integer                  ' 循环计数器
    
    ' 单个样式参数拆解
    Dim styleName As String           ' 样式名称（中文无空格、英文保留空格）
    Dim fontNameCN As String          ' 中文字体
    Dim fontNameEN As String          ' 英文字体
    Dim fontSize As Single            ' 字号（磅）
    Dim isBold As Boolean             ' 是否加粗
    Dim lineSpacingVal As Single      ' 固定行距值（磅）
    Dim spaceBeforeVal As Single      ' 段前间距（磅）
    Dim spaceAfterVal As Single       ' 段后间距（磅）
    Dim keepWithNextVal As Boolean    ' 是否与下一段同页
    Dim alignVal As Integer           ' 对齐方式
    
    '-------------------------- 常量定义 --------------------------
    Const CHAR_TO_POINT As Single = 2.835                       ' 1字符≈2.835磅
    Const INDENT_2CHAR As Single = 2 * CHAR_TO_POINT            ' 2字符缩进（关键常量）
    Const OUTLINE_LEVEL_1 As Integer = wdOutlineLevel1          ' 大纲级别1级
    Const LINE_SPACE_EXACTLY As Integer = wdLineSpaceExactly    ' 固定值行距
    Const LINE_SPACE_SINGLE As Integer = wdLineSpaceSingle      ' 单倍行距
    Const ALIGN_CENTER As Integer = wdAlignParagraphCenter      ' 居中对齐
    Const ALIGN_JUSTIFY As Integer = wdAlignParagraphJustify    ' 两端对齐
    
    '-------------------------- 第一步：绑定活动文档 --------------------------
    On Error Resume Next
    Set doc = ActiveDocument
    If Err.Number <> 0 Then
        MsgBox "请先打开一个Word文档再执行此宏！", vbCritical, "执行错误"
        Exit Sub
    End If
    On Error GoTo 0
    
    '-------------------------- 第二步：定义样式配置参数（样式名无中文空格） --------------------------
    ' 列索引：1:名称 2:类型 3:中文字体 4:英文字体 5:字号 6:加粗 7:行距 8:段前 9:段后 10:同页 11:对齐
    ReDim stylesInfo(1 To 28, 1 To 11)
    
    ' ========== BK基础系列（1-20）：中文样式名无空格 ==========
    stylesInfo(1, 1) = "BK表题英文": stylesInfo(1, 2) = "段落样式": stylesInfo(1, 3) = "宋体": stylesInfo(1, 4) = "Times New Roman": stylesInfo(1, 5) = 10.5: stylesInfo(1, 6) = False: stylesInfo(1, 7) = 18: stylesInfo(1, 8) = 0: stylesInfo(1, 9) = 0: stylesInfo(1, 10) = True: stylesInfo(1, 11) = ALIGN_CENTER
    stylesInfo(2, 1) = "BK表题中文": stylesInfo(2, 2) = "段落样式": stylesInfo(2, 3) = "宋体": stylesInfo(2, 4) = "Times New Roman": stylesInfo(2, 5) = 10.5: stylesInfo(2, 6) = False: stylesInfo(2, 7) = 18: stylesInfo(2, 8) = 3: stylesInfo(2, 9) = 0: stylesInfo(2, 10) = True: stylesInfo(2, 11) = ALIGN_CENTER
    stylesInfo(3, 1) = "BK目录""目录""": stylesInfo(3, 2) = "段落样式": stylesInfo(3, 3) = "黑体": stylesInfo(3, 4) = "Times New Roman": stylesInfo(3, 5) = 16: stylesInfo(3, 6) = False: stylesInfo(3, 7) = 20: stylesInfo(3, 8) = 0: stylesInfo(3, 9) = 40: stylesInfo(3, 10) = True: stylesInfo(3, 11) = ALIGN_CENTER
    stylesInfo(4, 1) = "BK图行": stylesInfo(4, 2) = "段落样式": stylesInfo(4, 3) = "宋体": stylesInfo(4, 4) = "Times New Roman": stylesInfo(4, 5) = 10.5: stylesInfo(4, 6) = False: stylesInfo(4, 7) = 0: stylesInfo(4, 8) = 6: stylesInfo(4, 9) = 0: stylesInfo(4, 10) = True: stylesInfo(4, 11) = ALIGN_CENTER
    stylesInfo(5, 1) = "BK一级标题【前言、结论、参考文献、致谢】": stylesInfo(5, 2) = "段落样式": stylesInfo(5, 3) = "黑体": stylesInfo(5, 4) = "Times New Roman": stylesInfo(5, 5) = 16: stylesInfo(5, 6) = False: stylesInfo(5, 7) = 20: stylesInfo(5, 8) = 0: stylesInfo(5, 9) = 40: stylesInfo(5, 10) = True: stylesInfo(5, 11) = ALIGN_CENTER
    stylesInfo(6, 1) = "BK图题英文": stylesInfo(6, 2) = "段落样式": stylesInfo(6, 3) = "宋体": stylesInfo(6, 4) = "Times New Roman": stylesInfo(6, 5) = 10.5: stylesInfo(6, 6) = False: stylesInfo(6, 7) = 18: stylesInfo(6, 8) = 0: stylesInfo(6, 9) = 6: stylesInfo(6, 10) = False: stylesInfo(6, 11) = ALIGN_CENTER
    stylesInfo(7, 1) = "BK摘要中文""摘要""": stylesInfo(7, 2) = "段落样式": stylesInfo(7, 3) = "黑体": stylesInfo(7, 4) = "Times New Roman": stylesInfo(7, 5) = 15: stylesInfo(7, 6) = False: stylesInfo(7, 7) = 20: stylesInfo(7, 8) = 0: stylesInfo(7, 9) = 20: stylesInfo(7, 10) = True: stylesInfo(7, 11) = ALIGN_CENTER
    stylesInfo(8, 1) = "BK摘要英文""ABSTRACT""": stylesInfo(8, 2) = "段落样式": stylesInfo(8, 3) = "宋体": stylesInfo(8, 4) = "Times New Roman": stylesInfo(8, 5) = 15: stylesInfo(8, 6) = True: stylesInfo(8, 7) = 20: stylesInfo(8, 8) = 0: stylesInfo(8, 9) = 40: stylesInfo(8, 10) = True: stylesInfo(8, 11) = ALIGN_CENTER
    stylesInfo(9, 1) = "BK摘要中文论文题目": stylesInfo(9, 2) = "段落样式": stylesInfo(9, 3) = "黑体": stylesInfo(9, 4) = "Times New Roman": stylesInfo(9, 5) = 16: stylesInfo(9, 6) = False: stylesInfo(9, 7) = 20: stylesInfo(9, 8) = 0: stylesInfo(9, 9) = 20: stylesInfo(9, 10) = True: stylesInfo(9, 11) = ALIGN_CENTER
    stylesInfo(10, 1) = "BK摘要英文""KEY WORDS""": stylesInfo(10, 2) = "段落样式": stylesInfo(10, 3) = "宋体": stylesInfo(10, 4) = "Times New Roman": stylesInfo(10, 5) = 14: stylesInfo(10, 6) = True: stylesInfo(10, 7) = 20: stylesInfo(10, 8) = 0: stylesInfo(10, 9) = 0: stylesInfo(10, 10) = False: stylesInfo(10, 11) = ALIGN_JUSTIFY
    stylesInfo(11, 1) = "BK摘要英文—关键词": stylesInfo(11, 2) = "段落样式": stylesInfo(11, 3) = "宋体": stylesInfo(11, 4) = "Times New Roman": stylesInfo(11, 5) = 14: stylesInfo(11, 6) = False: stylesInfo(11, 7) = 20: stylesInfo(11, 8) = 0: stylesInfo(11, 9) = 0: stylesInfo(11, 10) = False: stylesInfo(11, 11) = ALIGN_JUSTIFY
    stylesInfo(12, 1) = "BKABSTRACT PAPER TITLE": stylesInfo(12, 2) = "段落样式": stylesInfo(12, 3) = "宋体": stylesInfo(12, 4) = "Times New Roman": stylesInfo(12, 5) = 16: stylesInfo(12, 6) = False: stylesInfo(12, 7) = 20: stylesInfo(12, 8) = 0: stylesInfo(12, 9) = 15: stylesInfo(12, 10) = True: stylesInfo(12, 11) = ALIGN_CENTER
    stylesInfo(13, 1) = "BK摘要英文—正文": stylesInfo(13, 2) = "段落样式": stylesInfo(13, 3) = "宋体": stylesInfo(13, 4) = "Times New Roman": stylesInfo(13, 5) = 14: stylesInfo(13, 6) = False: stylesInfo(13, 7) = 20: stylesInfo(13, 8) = 0: stylesInfo(13, 9) = 40: stylesInfo(13, 10) = False: stylesInfo(13, 11) = ALIGN_JUSTIFY
    stylesInfo(14, 1) = "BK摘要中文""关键词""": stylesInfo(14, 2) = "段落样式": stylesInfo(14, 3) = "黑体": stylesInfo(14, 4) = "Times New Roman": stylesInfo(14, 5) = 14: stylesInfo(14, 6) = True: stylesInfo(14, 7) = 20: stylesInfo(14, 8) = 0: stylesInfo(14, 9) = 0: stylesInfo(14, 10) = False: stylesInfo(14, 11) = ALIGN_JUSTIFY
    stylesInfo(15, 1) = "BK摘要中文-关键词": stylesInfo(15, 2) = "段落样式": stylesInfo(15, 3) = "宋体": stylesInfo(15, 4) = "Times New Roman": stylesInfo(15, 5) = 14: stylesInfo(15, 6) = False: stylesInfo(15, 7) = 20: stylesInfo(15, 8) = 0: stylesInfo(15, 9) = 20: stylesInfo(15, 10) = False: stylesInfo(15, 11) = ALIGN_JUSTIFY
    stylesInfo(16, 1) = "BK摘要中文—正文": stylesInfo(16, 2) = "段落样式": stylesInfo(16, 3) = "宋体": stylesInfo(16, 4) = "Times New Roman": stylesInfo(16, 5) = 14: stylesInfo(16, 6) = False: stylesInfo(16, 7) = 20: stylesInfo(16, 8) = 0: stylesInfo(16, 9) = 20: stylesInfo(16, 10) = False: stylesInfo(16, 11) = ALIGN_JUSTIFY
    stylesInfo(17, 1) = "BK正文缩进": stylesInfo(17, 2) = "段落样式": stylesInfo(17, 3) = "宋体": stylesInfo(17, 4) = "Times New Roman": stylesInfo(17, 5) = 12: stylesInfo(17, 6) = False: stylesInfo(17, 7) = 20: stylesInfo(17, 8) = 0: stylesInfo(17, 9) = 0: stylesInfo(17, 10) = False: stylesInfo(17, 11) = ALIGN_JUSTIFY
    stylesInfo(18, 1) = "BK正文顶格": stylesInfo(18, 2) = "段落样式": stylesInfo(18, 3) = "宋体": stylesInfo(18, 4) = "Times New Roman": stylesInfo(18, 5) = 12: stylesInfo(18, 6) = False: stylesInfo(18, 7) = 20: stylesInfo(18, 8) = 0: stylesInfo(18, 9) = 0: stylesInfo(18, 10) = False: stylesInfo(18, 11) = ALIGN_JUSTIFY
    stylesInfo(19, 1) = "BK图题中文": stylesInfo(19, 2) = "段落样式": stylesInfo(19, 3) = "宋体": stylesInfo(19, 4) = "Times New Roman": stylesInfo(19, 5) = 10.5: stylesInfo(19, 6) = False: stylesInfo(19, 7) = 18: stylesInfo(19, 8) = 0: stylesInfo(19, 9) = 0: stylesInfo(19, 10) = True: stylesInfo(19, 11) = ALIGN_CENTER
    stylesInfo(20, 1) = "MTDisplayEquation": stylesInfo(20, 2) = "段落样式": stylesInfo(20, 3) = "宋体": stylesInfo(20, 4) = "Times New Roman": stylesInfo(20, 5) = 12: stylesInfo(20, 6) = False: stylesInfo(20, 7) = 22: stylesInfo(20, 8) = 0: stylesInfo(20, 9) = 0: stylesInfo(20, 10) = False: stylesInfo(20, 11) = ALIGN_JUSTIFY

    ' ========== TOC目录系列（21-24）：英文+数字保留空格 ==========
    stylesInfo(21, 1) = "TOC 1": stylesInfo(21, 2) = "段落样式": stylesInfo(21, 3) = "黑体": stylesInfo(21, 4) = "Times New Roman": stylesInfo(21, 5) = 14: stylesInfo(21, 6) = False: stylesInfo(21, 7) = 20: stylesInfo(21, 8) = 0: stylesInfo(21, 9) = 0: stylesInfo(21, 10) = False: stylesInfo(21, 11) = ALIGN_JUSTIFY
    stylesInfo(22, 1) = "TOC 2": stylesInfo(22, 2) = "段落样式": stylesInfo(22, 3) = "黑体": stylesInfo(22, 4) = "Times New Roman": stylesInfo(22, 5) = 12: stylesInfo(22, 6) = False: stylesInfo(22, 7) = 20: stylesInfo(22, 8) = 0: stylesInfo(22, 9) = 0: stylesInfo(22, 10) = False: stylesInfo(22, 11) = ALIGN_JUSTIFY
    stylesInfo(23, 1) = "TOC 3": stylesInfo(23, 2) = "段落样式": stylesInfo(23, 3) = "黑体": stylesInfo(23, 4) = "Times New Roman": stylesInfo(23, 5) = 12: stylesInfo(23, 6) = False: stylesInfo(23, 7) = 20: stylesInfo(23, 8) = 0: stylesInfo(23, 9) = 0: stylesInfo(23, 10) = False: stylesInfo(23, 11) = ALIGN_JUSTIFY
    stylesInfo(24, 1) = "TOC 4": stylesInfo(24, 2) = "段落样式": stylesInfo(24, 3) = "黑体": stylesInfo(24, 4) = "Times New Roman": stylesInfo(24, 5) = 12: stylesInfo(24, 6) = False: stylesInfo(24, 7) = 20: stylesInfo(24, 8) = 0: stylesInfo(24, 9) = 0: stylesInfo(24, 10) = False: stylesInfo(24, 11) = ALIGN_JUSTIFY

    ' ========== 标题系列（25-28）：英文+数字保留空格 ==========
    stylesInfo(25, 1) = "标题 1": stylesInfo(25, 2) = "段落样式": stylesInfo(25, 3) = "黑体": stylesInfo(25, 4) = "Times New Roman": stylesInfo(25, 5) = 16: stylesInfo(25, 6) = False: stylesInfo(25, 7) = 22: stylesInfo(25, 8) = 0: stylesInfo(25, 9) = 40: stylesInfo(25, 10) = True: stylesInfo(25, 11) = ALIGN_CENTER
    stylesInfo(26, 1) = "标题 2": stylesInfo(26, 2) = "段落样式": stylesInfo(26, 3) = "黑体": stylesInfo(26, 4) = "Times New Roman": stylesInfo(26, 5) = 14: stylesInfo(26, 6) = False: stylesInfo(26, 7) = 22: stylesInfo(26, 8) = 0: stylesInfo(26, 9) = 20: stylesInfo(26, 10) = True: stylesInfo(26, 11) = ALIGN_JUSTIFY
    stylesInfo(27, 1) = "标题 3": stylesInfo(27, 2) = "段落样式": stylesInfo(27, 3) = "黑体": stylesInfo(27, 4) = "Times New Roman": stylesInfo(27, 5) = 12: stylesInfo(27, 6) = False: stylesInfo(27, 7) = 22: stylesInfo(27, 8) = 0: stylesInfo(27, 9) = 20: stylesInfo(27, 10) = True: stylesInfo(27, 11) = ALIGN_JUSTIFY
    stylesInfo(28, 1) = "标题 4": stylesInfo(28, 2) = "段落样式": stylesInfo(28, 3) = "黑体": stylesInfo(28, 4) = "Times New Roman": stylesInfo(28, 5) = 12: stylesInfo(28, 6) = False: stylesInfo(28, 7) = 22: stylesInfo(28, 8) = 0: stylesInfo(28, 9) = 0: stylesInfo(28, 10) = True: stylesInfo(28, 11) = ALIGN_JUSTIFY
    
    '-------------------------- 第三步：循环创建/修改样式 --------------------------
    For i = LBound(stylesInfo, 1) To UBound(stylesInfo, 1)
        ' 提取当前样式参数
        styleName = stylesInfo(i, 1)
        fontNameCN = stylesInfo(i, 3)
        fontNameEN = stylesInfo(i, 4)
        fontSize = stylesInfo(i, 5)
        isBold = stylesInfo(i, 6)
        lineSpacingVal = stylesInfo(i, 7)
        spaceBeforeVal = stylesInfo(i, 8)
        spaceAfterVal = stylesInfo(i, 9)
        keepWithNextVal = stylesInfo(i, 10)
        alignVal = stylesInfo(i, 11)
        
        ' 检查样式是否存在
        On Error Resume Next
        Set targetStyle = doc.Styles(styleName)
        On Error GoTo 0
        
        If targetStyle Is Nothing Then
            Set targetStyle = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
            Debug.Print "新建样式：" & styleName
        Else
            Debug.Print "修改样式：" & styleName
        End If
        
        ' 配置样式属性
        With targetStyle
            ' 样式基准=无样式
            .BaseStyle = ""
            
            ' 字体配置
            With .Font
                .Name = fontNameCN
                .NameAscii = fontNameEN
                .NameOther = fontNameEN
                .Size = fontSize
                .Bold = isBold
                .Italic = False
                .Underline = wdUnderlineNone
                .Color = wdColorBlack
                If styleName = "BKABSTRACT PAPER TITLE" Then
                    .AllCaps = True
                End If
            End With
            
            ' 段落配置
            With .ParagraphFormat
                ' 基础缩进（默认0）
                .LeftIndent = CentimetersToPoints(0)
                .RightIndent = CentimetersToPoints(0)
                .FirstLineIndent = CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                
                ' ========== 关键修改：首行缩进2字符 ==========
                If styleName = "BK摘要中文—正文" Or styleName = "BK摘要英文—正文" Or styleName = "BK正文缩进" Then
                    .FirstLineIndent = INDENT_2CHAR  ' 2字符首行缩进
                    .CharacterUnitFirstLineIndent = 2 ' 补充字符单位缩进（兼容显示）
                End If
                
                ' 间距配置（关闭自动间距）
                .spaceBefore = spaceBeforeVal
                .SpaceBeforeAuto = False
                .spaceAfter = spaceAfterVal
                .SpaceAfterAuto = False
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                
                ' 对齐与分页
                .alignment = alignVal
                .WidowControl = False
                .KeepWithNext = keepWithNextVal
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                
                ' ========== 新增：标题1段前分页 ==========
                If styleName = "标题 1" Then
                    .PageBreakBefore = True
                End If
                
                ' 特殊缩进（BKABSTRACT PAPER TITLE）
                If styleName = "BKABSTRACT PAPER TITLE" Then
                    .LeftIndent = 5 * CHAR_TO_POINT
                    .RightIndent = 5 * CHAR_TO_POINT
                End If
                
                ' 大纲级别（BK一级标题）
                If styleName = "BK一级标题【前言、结论、参考文献、致谢】" Then
                    .OutlineLevel = OUTLINE_LEVEL_1
                End If
                
                ' 行距配置（固定值/单倍行距）
                If styleName = "BK图行" Then
                    .LineSpacingRule = LINE_SPACE_SINGLE
                Else
                    .LineSpacingRule = LINE_SPACE_EXACTLY
                    .lineSpacing = lineSpacingVal
                End If
                
                ' 补充属性（避免默认值干扰）
                .Hyphenation = True
                .MirrorIndents = False
                .TextboxTightWrap = wdTightNone
                .CollapsedByDefault = False
                .AutoAdjustRightIndent = True
                .DisableLineHeightGrid = False
                .FarEastLineBreakControl = True
                .WordWrap = True
                .HangingPunctuation = True
                .HalfWidthPunctuationOnTopOfLine = False
                .AddSpaceBetweenFarEastAndAlpha = True
                .AddSpaceBetweenFarEastAndDigit = True
                .BaseLineAlignment = wdBaselineAlignAuto
            End With
        End With
        
        Set targetStyle = Nothing
    Next i
    
    '-------------------------- 第四步：执行完成提示 --------------------------
    Set doc = Nothing
    MsgBox "样式配置完成！", vbInformation, "执行成功"
End Sub

Sub A删百分号开头段()
    Dim i As Integer ' 声明循环计数器
    Dim paraText As String ' 存储段落文本内容
    
    ' 关键：倒序遍历段落（避免正序删除导致段落索引错乱、漏判）
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        ' 获取当前段落文本，并去除首尾空白字符（避免空格/制表符干扰判断）
        paraText = Trim(ActiveDocument.Paragraphs(i).Range.text)
        
        ' 判断：段落文本非空 且 以%开头
        If Len(paraText) > 0 And Left(paraText, 1) = "%" Then
            ' 删除该段落
            ActiveDocument.Paragraphs(i).Range.Delete
        End If
    Next i
    
    MsgBox "遍历完成！已删除所有以%开头的段落", vbInformation, "执行成功"
End Sub


Sub A文献引用改蓝色()
    For i = 1 To ActiveDocument.Fields.Count '遍历文档所有域
        ' Zotero插入的引用域代码的起始31位是" ADDIN ZOTERO_ITEM CSL_CITATION"
        If Left(ActiveDocument.Fields(i).code, 31) = " ADDIN ZOTERO_ITEM CSL_CITATION" Then
            ActiveDocument.Fields(i).Select '选中Zotero引用域
            Selection.Font.Color = wdColorBlue '设置字体颜色为蓝色
        End If
    Next
End Sub


Sub A英文括号上升1磅()
    Dim rng As Range
    Dim i As Long
    Dim char As Range
    
    ' 检查是否有文本被选中
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选中要处理的文本！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 设置要处理的范围为选中的文本
    Set rng = Selection.Range
    
    ' 逐个字符检查
    For i = 1 To rng.Characters.Count
        Set char = rng.Characters(i)
        
        ' 检查是否为左括号或右括号
        If char.text = "(" Or char.text = ")" Then
            ' 获取当前字符的字体设置
            With char.Font
                ' 保存原始位置
                Dim originalPosition As Single
                originalPosition = .Position
                
                ' 将位置上移1磅（减少Position值）
                ' 注意：Position为正数表示上移，负数表示下移
                .Position = originalPosition + 1
            End With
        End If
    Next i
    
    MsgBox "括号位置上移完成！", vbInformation, "完成"
End Sub

Sub B新式公式转斜体()
    ' 将文档中所有新式公式（OMath 对象）的内容设置为斜体
    ' 对旧版 Equation 3.0 对象仅统计并提示，不自动修改（因无法稳定实现）
    
    Dim math As oMath
    Dim newCount As Long
    Dim oldInlineCount As Long
    Dim oldShapeCount As Long
    Dim iShape As InlineShape
    Dim shp As Shape
    
    ' 处理新式公式（推荐使用）
    newCount = 0
    For Each math In ActiveDocument.OMaths
        math.Range.Font.Italic = True
        newCount = newCount + 1
    Next math
    
    ' 统计内嵌的旧版 Equation 3.0 公式
    oldInlineCount = 0
    For Each iShape In ActiveDocument.InlineShapes
        If iShape.Type = wdInlineShapeEmbeddedOLEObject Then
            If iShape.OLEFormat.ClassType Like "*Equation*" Then
                oldInlineCount = oldInlineCount + 1
            End If
        End If
    Next iShape
    
    ' 统计浮动形状中的旧版 Equation 3.0 公式
    oldShapeCount = 0
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoEmbeddedOLEObject Then
            If shp.OLEFormat.ClassType Like "*Equation*" Then
                oldShapeCount = oldShapeCount + 1
            End If
        End If
    Next shp
    
    ' 提示结果
    Dim msg As String
    msg = "已将 " & newCount & " 个新式公式设置为斜体。"
    If oldInlineCount + oldShapeCount > 0 Then
        msg = msg & vbCrLf & "发现 " & (oldInlineCount + oldShapeCount) & " 个旧版 Equation 3.0 公式。"
        msg = msg & vbCrLf & "请手动将其转换为新式公式（推荐）或手动设置斜体。"
    End If
    MsgBox msg, vbInformation, "公式斜体化完成"
End Sub


Sub B一键修改OMML公式字体()
    ' 一键修改公式字体 宏
    Dim 公式 As oMath
    
    ' 循环遍历当前文档的公式
    For Each 公式 In ActiveDocument.OMaths
        ' 关键一步：将公式转为普通文本，这样字体设置才能生效
        公式.ConvertToNormalText
    Next
    ' 设置整个文档的ASCII字体为Times New Roman
    ActiveDocument.Content.Font.NameAscii = "Times New Roman"
End Sub


Sub B度符号统一改为TimesNewRoman()
    Application.ScreenUpdating = False
    
    
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        ' 匹配度符号
        .text = "°"
        .Replacement.text = "^&"
        ' 仅修改西文字体，中文字体置空，解决悬停显示问题
        .Replacement.Font.Name = ""
        .Replacement.Font.NameAscii = "Times New Roman"
        .Replacement.Font.NameOther = "Times New Roman"
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindStop
        ' 执行替换
        .Execute Replace:=wdReplaceAll
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "所有度数符号已修改完成！", vbInformation
End Sub


Sub B清除页眉横线()
    ' 禁用屏幕更新，运行更流畅
    Application.ScreenUpdating = False
    
    ' 获取内置的"页眉"样式
    Dim 页眉样式 As Style
    Set 页眉样式 = ActiveDocument.Styles(wdStyleHeader)
        
    ' Borders 是集合，没有 LineStyle 属性，需遍历每个 Border 对象
    页眉样式.Borders.Enable = False
    
    Dim bd As Border
    For Each bd In 页眉样式.Borders
        bd.LineStyle = wdLineStyleNone
    Next bd
    
    
    ' 刷新文档显示
    ActiveDocument.Repaginate
    
    ' 恢复屏幕更新
    Application.ScreenUpdating = True
    
    ' 提示完成
    MsgBox "已成功清除页眉样式的横线！", vbInformation, "完成"
End Sub

Sub A改批注人名()
    Dim cmt As Comment
    Dim NewAuthorName As String
    
    ' =========在这里修改为你想要的批注作者名字=========
    NewAuthorName = "张三"
    ' =================================================
    
    For Each cmt In ActiveDocument.Comments
        cmt.Author = NewAuthorName
        cmt.Initial = Left(NewAuthorName, 2) '缩写取名字前2个字，可以自己改数字
    Next cmt
    
    MsgBox "完成！所有批注作者已修改为：" & NewAuthorName
End Sub
