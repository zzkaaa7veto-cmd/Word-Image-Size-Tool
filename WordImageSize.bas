Attribute VB_Name = "WordImageSize"
Option Explicit

' All measurements below are in centimeters.
Private Const DEFAULT_WIDTH_CM As Double = 5#

Private Sub ResizeInline(ByVal pic As InlineShape, ByVal widthCm As Double, _
                         ByVal keepRatio As Boolean, ByVal heightCm As Double)
    If pic.Type <> wdInlineShapePicture And pic.Type <> wdInlineShapeLinkedPicture Then Exit Sub
    pic.LockAspectRatio = IIf(keepRatio, msoTrue, msoFalse)
    pic.Width = CentimetersToPoints(widthCm)
    If Not keepRatio Then pic.Height = CentimetersToPoints(heightCm)
End Sub

Private Sub ResizeFloating(ByVal pic As Shape, ByVal widthCm As Double, _
                           ByVal keepRatio As Boolean, ByVal heightCm As Double)
    If pic.Type <> msoPicture And pic.Type <> msoLinkedPicture Then Exit Sub
    pic.LockAspectRatio = IIf(keepRatio, msoTrue, msoFalse)
    pic.Width = CentimetersToPoints(widthCm)
    If Not keepRatio Then pic.Height = CentimetersToPoints(heightCm)
End Sub

Private Function AskDimension(ByVal title As String, ByVal defaultValue As Double) As Double
    Dim answer As String
    answer = InputBox("请输入厘米数（必须大于 0）：", title, CStr(defaultValue))
    If Len(Trim$(answer)) = 0 Then Exit Function
    If Not IsNumeric(answer) Then
        MsgBox "请输入有效的数字。", vbExclamation
        Exit Function
    End If
    AskDimension = CDbl(answer)
    If AskDimension <= 0# Then
        AskDimension = 0#
        MsgBox "尺寸必须大于 0。", vbExclamation
    End If
End Function

' 1. Resize all ordinary inline and floating pictures, proportional.
Public Sub ResizeAllPicturesProportionally()
    Dim w As Double, pic As InlineShape, shp As Shape, count As Long
    w = AskDimension("全部图片统一宽度", DEFAULT_WIDTH_CM)
    If w <= 0 Then Exit Sub
    For Each pic In ActiveDocument.InlineShapes
        If pic.Type = wdInlineShapePicture Or pic.Type = wdInlineShapeLinkedPicture Then
            ResizeInline pic, w, True, 0
            count = count + 1
        End If
    Next pic
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ResizeFloating shp, w, True, 0
            count = count + 1
        End If
    Next shp
    MsgBox "已处理 " & count & " 张图片。", vbInformation
End Sub

' 2. Resize pictures inside the selected text range; also supports Ctrl-click
' selected floating pictures if Word exposes them as a ShapeRange.
Public Sub ResizeSelectedPicturesProportionally()
    Dim w As Double, pic As InlineShape, shp As Shape, count As Long
    Dim shapeRange As ShapeRange, i As Long
    w = AskDimension("所选图片统一宽度", DEFAULT_WIDTH_CM)
    If w <= 0 Then Exit Sub
    On Error Resume Next
    Set shapeRange = Selection.ShapeRange
    On Error GoTo 0
    If Not shapeRange Is Nothing Then
        For i = 1 To shapeRange.Count
            Set shp = shapeRange(i)
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                ResizeFloating shp, w, True, 0
                count = count + 1
            End If
        Next i
    Else
        For Each pic In Selection.Range.InlineShapes
            If pic.Type = wdInlineShapePicture Or pic.Type = wdInlineShapeLinkedPicture Then
                ResizeInline pic, w, True, 0
                count = count + 1
            End If
        Next pic
    End If
    MsgBox "已处理所选范围内 " & count & " 张图片。" & vbCrLf & _
           "不连续的嵌入式图片请使用按编号处理功能。", vbInformation
End Sub

' 3. List pictures with sequential indices in document order by class.
' The two numbering systems are separate: I1/I2... (inline) and F1/F2... (floating).
Public Sub ListPictureIndices()
    Dim i As Long, msg As String
    msg = "嵌入式图片（I）：" & vbCrLf
    For i = 1 To ActiveDocument.InlineShapes.Count
        msg = msg & "I" & i & "  " & ActiveDocument.InlineShapes(i).Range.Information(wdActiveEndPageNumber) & "页" & vbCrLf
    Next i
    msg = msg & vbCrLf & "浮动图片（F，按锚点所在页）：" & vbCrLf
    For i = 1 To ActiveDocument.Shapes.Count
        msg = msg & "F" & i & "  " & ActiveDocument.Shapes(i).Anchor.Information(wdActiveEndPageNumber) & "页" & vbCrLf
    Next i
    Debug.Print msg
    MsgBox "图片索引列表已写入 VBA 立即窗口（Ctrl+G 查看）。" & vbCrLf & _
           "嵌入式：I1、I2；浮动式：F1、F2。", vbInformation
End Sub

' 4. Resize non-adjacent images by typed indices, e.g. I1,I3,F2.
Public Sub ResizePicturesByIndices()
    Dim raw As String, parts As Variant, item As Variant
    Dim token As String, kind As String, n As Long, w As Double, done As Long, skipped As Long
    raw = InputBox("输入图片编号，以英文逗号分隔，例如 I1,I3,F2。" & vbCrLf & _
                   "先运行 ListPictureIndices 查看编号。", "按编号指定图片")
    If Len(Trim$(raw)) = 0 Then Exit Sub
    w = AskDimension("指定图片统一宽度", DEFAULT_WIDTH_CM)
    If w <= 0 Then Exit Sub
    parts = Split(UCase$(Replace(raw, " ", "")), ",")
    For Each item In parts
        token = CStr(item)
        If Len(token) < 2 Then
            skipped = skipped + 1
        Else
            kind = Left$(token, 1)
            If Not IsNumeric(Mid$(token, 2)) Then
                skipped = skipped + 1
            Else
                n = CLng(Mid$(token, 2))
                If kind = "I" And n >= 1 And n <= ActiveDocument.InlineShapes.Count Then
                    ResizeInline ActiveDocument.InlineShapes(n), w, True, 0
                    done = done + 1
                ElseIf kind = "F" And n >= 1 And n <= ActiveDocument.Shapes.Count Then
                    ResizeFloating ActiveDocument.Shapes(n), w, True, 0
                    done = done + 1
                Else
                    skipped = skipped + 1
                End If
            End If
        End If
    Next item
    MsgBox "处理 " & done & " 项，忽略无效编号 " & skipped & " 项。", vbInformation
End Sub

' 5. Exact size (may distort photos); applies only to selected pictures.
Public Sub ResizeSelectedPicturesExactSize()
    Dim w As Double, h As Double, pic As InlineShape, shp As Shape
    Dim sr As ShapeRange, i As Long, count As Long
    w = AskDimension("指定宽度", 5#)
    If w <= 0 Then Exit Sub
    h = AskDimension("指定高度", 3#)
    If h <= 0 Then Exit Sub
    On Error Resume Next
    Set sr = Selection.ShapeRange
    On Error GoTo 0
    If Not sr Is Nothing Then
        For i = 1 To sr.Count
            Set shp = sr(i)
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                ResizeFloating shp, w, False, h
                count = count + 1
            End If
        Next i
    Else
        For Each pic In Selection.Range.InlineShapes
            If pic.Type = wdInlineShapePicture Or pic.Type = wdInlineShapeLinkedPicture Then
                ResizeInline pic, w, False, h
                count = count + 1
            End If
        Next pic
    End If
    MsgBox "已设置 " & count & " 张图片为 " & w & " × " & h & " cm。", vbInformation
End Sub
