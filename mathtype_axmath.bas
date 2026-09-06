Attribute VB_Name = "Md2DocxTools"
'==========================================================================
' md2docx 配套 Word 宏（Windows + MS Word）
'
' 装法：Word 里按 Alt+F11 -> 文件 -> 导入文件 -> 选本文件 -> 回到 Word 按 Alt+F8 运行
'
' 可运行的宏：
'   UpdateAllFields        更新全文域：公式 SEQ 编号 + 正文 REF 引用联动。完全可靠。
'   SelectLatexEquations   选中/标记全部待转换的 LaTeX 源码，便于用 GUI 批量转换
'   ProbeEquationMacros    探测本机 MathType / AxMath 到底暴露了哪些可调用的宏
'   ConvertLatexToMathType 逐个把 LaTeX 源码转成 MathType 公式（依赖探测结果）
'   ConvertLatexToAxMath   同上，走 AxMath
'   CountLatexEquations    统计还剩多少个未转换的 LaTeX 源码
'   CleanupAfterConvert    转换后清理：去掉残留的 MathSource 样式与 $ 定界符
'
' 重要说明：
'   MathType / AxMath 都没有公开稳定的 VBA 接口文档。目前能查到的唯一线索是
'   MathType 6.5 时代的 MTCommand_TeXToggle，现在是 MathType 7，名字是否还在
'   无法确认。所以这里不硬编码猜测的名字，而是先探测一遍再用；探测不到就明确
'   提示改走 GUI 批量转换（那条路是官方文档写明的，可靠）。
'   请先运行 ProbeEquationMacros 看本机情况。
'==========================================================================
Option Explicit

' md2docx 给 LaTeX 源码打的字符样式名，与 make_reference.py 里的 MATHSRC_STYLE_ID 一致
Private Const MATHSRC_STYLE As String = "MathSource"

' 候选宏名。探测通过后，把确认可用的那个填到下面两个常量里可跳过探测。
Private Const MATHTYPE_MACRO_OVERRIDE As String = ""
Private Const AXMATH_MACRO_OVERRIDE As String = ""


'--------------------------------------------------------------------------
' 更新全文域。SEQ 编号和 REF 引用互相依赖，需要更新两遍才能稳定。
' 这个宏不依赖任何第三方插件，永远可用。
'--------------------------------------------------------------------------
Public Sub UpdateAllFields()
    Dim rng As Range, i As Integer, story As Range
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
    Application.ScreenUpdating = True
    MsgBox "已更新全文域两遍（公式 SEQ 编号 + 正文 REF 引用）。", vbInformation
End Sub


'--------------------------------------------------------------------------
' 统计还剩多少个未转换的 LaTeX 源码
'--------------------------------------------------------------------------
Public Sub CountLatexEquations()
    Dim n As Long
    n = CountRemaining()
    MsgBox "当前还有 " & n & " 处带 " & MATHSRC_STYLE & " 样式的 LaTeX 源码待转换。", vbInformation
End Sub

Private Function CountRemaining() As Long
    Dim n As Long
    Dim r As Range
    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .Text = ""
        .Style = MATHSRC_STYLE
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        Do While .Execute
            n = n + 1
            If r.End >= ActiveDocument.Content.End Then Exit Do
            r.SetRange r.End, ActiveDocument.Content.End
        Loop
    End With
    CountRemaining = n
End Function


'--------------------------------------------------------------------------
' 把全部 LaTeX 源码高亮出来，便于肉眼确认范围，或框选后用 GUI 批量转换。
'--------------------------------------------------------------------------
Public Sub SelectLatexEquations()
    Dim r As Range, n As Long
    Application.ScreenUpdating = False
    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .Text = ""
        .Style = MATHSRC_STYLE
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Replacement.Text = ""
        Do While .Execute
            r.HighlightColorIndex = wdYellow
            n = n + 1
            If r.End >= ActiveDocument.Content.End Then Exit Do
            r.SetRange r.End, ActiveDocument.Content.End
        Loop
    End With
    Application.ScreenUpdating = True
    MsgBox "已把 " & n & " 处 LaTeX 源码标为黄色高亮。" & vbCrLf & vbCrLf & _
           "接下来可以：" & vbCrLf & _
           "  MathType：MathType 选项卡 -> Convert Equations" & vbCrLf & _
           "            输入勾 Text using a translator，选 TeX 翻译器，范围 Whole document" & vbCrLf & _
           "  AxMath  ：用插件菜单的批量处理功能" & vbCrLf & vbCrLf & _
           "转换完运行 CleanupAfterConvert 清掉高亮和残留样式。", vbInformation
End Sub


'--------------------------------------------------------------------------
' 探测本机可用的公式插件宏。不改动文档，只报告。
'--------------------------------------------------------------------------
Public Sub ProbeEquationMacros()
    Dim cands As Variant, i As Long, msg As String, okList As String
    cands = AllCandidates()
    msg = "已安装的 Word 加载项：" & vbCrLf & ListAddins() & vbCrLf & _
          "宏名探测结果：" & vbCrLf

    For i = LBound(cands) To UBound(cands)
        If MacroCallable(CStr(cands(i))) Then
            msg = msg & "  [可调用] " & cands(i) & vbCrLf
            okList = okList & cands(i) & "; "
        Else
            msg = msg & "  [ 不可用 ] " & cands(i) & vbCrLf
        End If
    Next i

    If Len(okList) = 0 Then
        msg = msg & vbCrLf & _
              "没有探测到可调用的宏。这不代表插件没装，只说明它没暴露 VBA 接口。" & vbCrLf & _
              "请改走 GUI 批量转换：运行 SelectLatexEquations 看清范围，然后用" & vbCrLf & _
              "MathType 的 Convert Equations 或 AxMath 的批量处理功能。"
    Else
        msg = msg & vbCrLf & "可用：" & okList & vbCrLf & _
              "把它填进本模块顶部的 MATHTYPE_MACRO_OVERRIDE / AXMATH_MACRO_OVERRIDE 常量，" & vbCrLf & _
              "之后 ConvertLatexToMathType / ConvertLatexToAxMath 就会直接用它。"
    End If
    MsgBox msg, vbInformation, "公式插件探测"
End Sub

Private Function AllCandidates() As Variant
    ' MathType 与 AxMath 历史上出现过的宏名，从最可能的排起
    AllCandidates = Array( _
        "MTCommand_TeXToggle", _
        "MathType.MTCommand_TeXToggle", _
        "MathTypeCommands.MTCommand_TeXToggle", _
        "MTCommand_ConvertEqns", _
        "MathType.MTCommand_ConvertEqns", _
        "AxMath.ConvertLatex", _
        "AxMathCommands.ConvertLatex", _
        "AxMath.LatexToEquation")
End Function

Private Function MacroCallable(ByVal macroName As String) As Boolean
    ' 用一个不可能的参数去调，靠错误号区分「宏不存在」和「宏存在但参数不对」
    On Error Resume Next
    Err.Clear
    Application.Run macroName
    ' 5 / 待用宏不存在时 Word 一般报 429 或 5111；能跑通或报参数错都说明名字存在
    MacroCallable = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ListAddins() As String
    Dim ad As AddIn, s As String
    On Error Resume Next
    For Each ad In Application.AddIns
        s = s & "  " & ad.Name & IIf(ad.Installed, " [已启用]", " [未启用]") & vbCrLf
    Next ad
    If Len(s) = 0 Then s = "  (读取不到)" & vbCrLf
    ListAddins = s
End Function


'--------------------------------------------------------------------------
' 批量转换。逐个选中 LaTeX 源码后调用插件宏。
'--------------------------------------------------------------------------
Public Sub ConvertLatexToMathType()
    ConvertWith ResolveMacro(MATHTYPE_MACRO_OVERRIDE, "MTCommand"), "MathType"
End Sub

Public Sub ConvertLatexToAxMath()
    ConvertWith ResolveMacro(AXMATH_MACRO_OVERRIDE, "AxMath"), "AxMath"
End Sub

Private Function ResolveMacro(ByVal override As String, ByVal hint As String) As String
    Dim cands As Variant, i As Long
    If Len(override) > 0 Then
        ResolveMacro = override
        Exit Function
    End If
    cands = AllCandidates()
    For i = LBound(cands) To UBound(cands)
        If InStr(1, CStr(cands(i)), hint, vbTextCompare) > 0 Then
            If MacroCallable(CStr(cands(i))) Then
                ResolveMacro = CStr(cands(i))
                Exit Function
            End If
        End If
    Next i
    ResolveMacro = ""
End Function

Private Sub ConvertWith(ByVal macroName As String, ByVal label As String)
    Dim r As Range, okN As Long, failN As Long, total As Long

    If Len(macroName) = 0 Then
        MsgBox "没找到 " & label & " 可调用的 VBA 宏。" & vbCrLf & vbCrLf & _
               "请先运行 ProbeEquationMacros 看探测结果；" & vbCrLf & _
               "如果确实没有 VBA 接口，就走 GUI 批量转换：" & vbCrLf & _
               "  1. 运行 SelectLatexEquations 把待转换内容标出来" & vbCrLf & _
               "  2. " & label & " 的批量转换功能（MathType 是 Convert Equations，" & vbCrLf & _
               "     输入选 Text using a translator -> TeX，范围 Whole document）", _
               vbExclamation
        Exit Sub
    End If

    total = CountRemaining()
    If total = 0 Then
        MsgBox "没有待转换的 LaTeX 源码。文档可能已经转过，或者不是用 --math-mode latex 生成的。", vbInformation
        Exit Sub
    End If
    If MsgBox("准备用 " & label & "（宏：" & macroName & "）转换 " & total & _
              " 处公式。建议先另存一份备份。继续？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    ' 从后往前找：转换会改变文档长度，倒序处理不会让位置失效
    Do
        Set r = FindLastMathSource()
        If r Is Nothing Then Exit Do
        r.Select
        On Error Resume Next
        Err.Clear
        Application.Run macroName
        If Err.Number = 0 Then
            okN = okN + 1
        Else
            failN = failN + 1
            ' 失败就把样式清掉，避免下一轮重复找到同一处造成死循环
            Selection.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        End If
        On Error GoTo 0
    Loop
    Application.ScreenUpdating = True

    MsgBox label & " 转换完成：成功 " & okN & " 处，失败 " & failN & " 处。" & vbCrLf & vbCrLf & _
           "请务必抽查符号是否正确——TeX/OMML 转 " & label & " 已知会漏掉一些形近异码符号。" & vbCrLf & _
           "然后运行 UpdateAllFields 刷新公式编号与引用。", vbInformation
End Sub

Private Function FindLastMathSource() As Range
    Dim r As Range
    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .Text = ""
        .Style = MATHSRC_STYLE
        .Forward = False          ' 倒序
        .Wrap = wdFindStop
        .Format = True
        If .Execute Then
            Set FindLastMathSource = r
        Else
            Set FindLastMathSource = Nothing
        End If
    End With
End Function


'--------------------------------------------------------------------------
' 转换后清理：去高亮、清残留字符样式、去掉可能剩下的 $ 定界符
'--------------------------------------------------------------------------
Public Sub CleanupAfterConvert()
    Dim n As Long
    Application.ScreenUpdating = False

    ' 去黄色高亮
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Highlight = True
        .Replacement.Highlight = False
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With

    ' 去掉遗留的 $$ 和 $（只在 MathSource 样式范围内动手，避免误伤正文里的美元符号）
    n = ReplaceInStyle("$$", "")
    n = n + ReplaceInStyle("$", "")

    Application.ScreenUpdating = True
    MsgBox "清理完成。去掉了 " & n & " 处残留的 $ 定界符。" & vbCrLf & _
           "仍带 " & MATHSRC_STYLE & " 样式的位置还有 " & CountRemaining() & " 处。", vbInformation
End Sub

Private Function ReplaceInStyle(ByVal findText As String, ByVal replText As String) As Long
    Dim cnt As Long, r As Range
    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Style = MATHSRC_STYLE
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceOne)
            cnt = cnt + 1
            If cnt > 5000 Then Exit Do
        Loop
    End With
    ReplaceInStyle = cnt
End Function
