Attribute VB_Name = "Md2DocxTools"
' Windows + Microsoft Word。导入本模块后可更新域、标记源码及调用已确认的插件宏。
' ProbeEquationMacros 只报告加载项和配置，不试执行宏。
Option Explicit

Private Const MATHSRC_STYLE As String = "MathSource"
' 仅填写在本机插件文档或受控试运行中确认的宏名；空值时使用插件 GUI。
Private Const MATHTYPE_MACRO_OVERRIDE As String = ""
Private Const AXMATH_MACRO_OVERRIDE As String = ""

Public Sub UpdateAllFields()
    Dim story As Range, rng As Range, i As Integer, updating As Boolean
    updating = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    For i = 1 To 2
        For Each story In ActiveDocument.StoryRanges
            Set rng = story
            Do
                rng.Fields.Update
                Set rng = rng.NextStoryRange
            Loop Until rng Is Nothing
        Next story
    Next i
    Application.ScreenUpdating = updating
    MsgBox "已更新全文域两遍。", vbInformation
    Exit Sub
Failed:
    Application.ScreenUpdating = updating
    MsgBox "更新域失败：" & Err.Description, vbExclamation
End Sub

' 统一查找范围：正文、脚注、尾注、页眉页脚等所有 story 及其后续范围。
' 先保存 Range 副本，转换时倒序遍历，每个源码只尝试一次，失败仍保留源码样式。
Private Function StyledRanges(ByVal doc As Document, Optional ByVal pendingOnly As Boolean = True) As Collection
    Dim result As New Collection
    Dim story As Range, part As Range, search As Range
    For Each story In doc.StoryRanges
        Set part = story
        Do
            Set search = part.Duplicate
            With search.Find
                .ClearFormatting
                .Text = ""
                .Style = MATHSRC_STYLE
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchCase = False
                Do While .Execute
                    If Not pendingOnly Or InStr(search.Text, "$") > 0 Then
                        result.Add search.Duplicate
                    End If
                    If search.End >= part.End Then Exit Do
                    search.SetRange search.End, part.End
                Loop
            End With
            Set part = part.NextStoryRange
        Loop Until part Is Nothing
    Next story
    Set StyledRanges = result
End Function

Private Function CountPending(ByVal doc As Document) As Long
    Dim sources As Collection
    Set sources = StyledRanges(doc)
    CountPending = sources.Count
End Function

Public Sub CountLatexEquations()
    MsgBox "当前还有 " & CountPending(ActiveDocument) & " 处 LaTeX 源码待转换（含脚注、尾注）。", vbInformation
End Sub

Public Sub SelectLatexEquations()
    Dim sources As Collection, r As Range
    Set sources = StyledRanges(ActiveDocument)
    For Each r In sources
        r.HighlightColorIndex = wdYellow
    Next r
    MsgBox "已高亮 " & sources.Count & " 处源码（含脚注、尾注）。请用插件 GUI 批量转换，或配置已确认的宏名。", vbInformation
End Sub

' 保留旧入口名称，但不再把执行候选宏当作探测。
Public Sub ProbeEquationMacros()
    Dim msg As String
    msg = "已安装的 Word 加载项：" & vbCrLf & ListAddins() & vbCrLf & _
          "MathType 配置：" & ConfiguredName(MATHTYPE_MACRO_OVERRIDE) & vbCrLf & _
          "AxMath 配置：" & ConfiguredName(AXMATH_MACRO_OVERRIDE) & vbCrLf & vbCrLf & _
          "此报告未执行任何插件宏，也不验证宏是否可调用。" & vbCrLf & _
          "没有已确认宏名时，请用插件 GUI 批量转换；不要通过试执行猜测接口。"
    MsgBox msg, vbInformation, "公式插件配置"
End Sub

Private Function ConfiguredName(ByVal macroName As String) As String
    If Len(macroName) = 0 Then
        ConfiguredName = "未配置（使用插件 GUI）"
    Else
        ConfiguredName = macroName & "（尚未在此报告中验证）"
    End If
End Function

Private Function ListAddins() As String
    Dim ad As AddIn, s As String
    On Error GoTo Failed
    For Each ad In Application.AddIns
        s = s & ad.Name & IIf(ad.Installed, " [已启用]", " [未启用]") & vbCrLf
    Next ad
    If Len(s) = 0 Then s = "(没有加载项)" & vbCrLf
    ListAddins = s
    Exit Function
Failed:
    ListAddins = "(无法读取加载项：" & Err.Description & ")" & vbCrLf
End Function

Public Sub ConvertLatexToMathType()
    ConvertWith MATHTYPE_MACRO_OVERRIDE, "MathType"
End Sub

Public Sub ConvertLatexToAxMath()
    ConvertWith AXMATH_MACRO_OVERRIDE, "AxMath"
End Sub

Private Sub ConvertWith(ByVal macroName As String, ByVal label As String)
    Dim doc As Document, total As Long, okN As Long, failN As Long, updating As Boolean
    Set doc = ActiveDocument
    updating = Application.ScreenUpdating
    On Error GoTo Failed
    If Len(macroName) = 0 Then
        MsgBox "未配置已确认的 " & label & " 宏名。请使用插件 GUI，或在模块顶部配置本机已验证的宏名。", vbExclamation
        Exit Sub
    End If
    total = CountPending(doc)
    If total = 0 Then
        MsgBox "没有待转换的 LaTeX 源码（已检查正文、脚注和尾注）。", vbInformation
        Exit Sub
    End If
    If MsgBox("准备用 " & label & "（" & macroName & "）转换 " & total & _
              " 处公式。请先另存备份。继续？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    ConvertPending doc, macroName, okN, failN
    Application.ScreenUpdating = updating
    MsgBox "已处理：成功 " & okN & "，失败或未消耗源码 " & failN & _
           "，剩余 " & CountPending(doc) & "。失败源码保留，可重试；请抽查公式并更新域。", vbInformation
    Exit Sub
Failed:
    Application.ScreenUpdating = updating
    MsgBox "转换中断：" & Err.Description & "。未转换源码仍保留。", vbExclamation
End Sub

Private Sub ConvertPending(ByVal doc As Document, ByVal macroName As String, ByRef okN As Long, ByRef failN As Long)
    Dim sources As Collection, r As Range, i As Long, before As Long, after As Long, errorCode As Long
    Set sources = StyledRanges(doc)
    For i = sources.Count To 1 Step -1
        Set r = sources(i)
        ' 某些插件会一次转换多处；跳过已被前一次调用消耗的快照，避免再切换回源码。
        If InStr(r.Text, "$") > 0 Then
            before = CountPending(doc)
            doc.Activate
            r.Select
            On Error Resume Next
            Err.Clear
            Application.Run macroName
            errorCode = Err.Number
            On Error GoTo 0
            after = CountPending(doc)
            If errorCode = 0 And after < before Then
                okN = okN + before - after
            Else
                failN = failN + 1
            End If
        End If
    Next i
End Sub

Public Sub CleanupAfterConvert()
    Dim ranges As Collection, r As Range
    Set ranges = StyledRanges(ActiveDocument, False)
    For Each r In ranges
        ' 未转换源码仍需 $ 定界符供后续重试；不清掉其样式或高亮。
        If InStr(r.Text, "$") = 0 Then
            r.HighlightColorIndex = wdNoHighlight
            r.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        End If
    Next r
    MsgBox "已清理转换后残留样式。仍待转换 " & CountPending(ActiveDocument) & _
           " 处；其源码和 $ 定界符保持不变。", vbInformation
End Sub
