# 可选实机回归：Windows + Word，并由用户预先启用对 VBA 工程对象模型的信任。
# 不修改信任设置、不依赖 MathType/AxMath；在临时文档中用受控宏替身验证 Word 范围行为。
# powershell -File example/test_word_macros.ps1
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $module = $doc.VBProject.VBComponents.Add(1)
    $module.Name = 'Md2DocxTools'
    $source = Get-Content (Join-Path $PSScriptRoot '../mathtype_axmath.bas') -Raw -Encoding UTF8
    $source = $source -replace '(?m)^Attribute[^\r\n]*\r?\n', ''
    # 仅替换测试副本中的 UI 对话框；文档查找和插件调用路径保持原样。
    $source = $source -replace '\bMsgBox\b', 'TestMessage'
    $source = $source.Replace('Option Explicit', "Option Explicit`r`nPrivate TestCalls As Long`r`nPrivate TestAnswer As VbMsgBoxResult")
    $tests = @'

Private Function TestMessage(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    TestMessage = TestAnswer
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 513, "VBA regression", message
End Sub

Public Sub TestConvertSource()
    TestCalls = TestCalls + 1
    Selection.Range.Text = "converted"
End Sub

Public Sub TestNoOp()
    TestCalls = TestCalls + 1
End Sub

Public Sub TestFailure()
    TestCalls = TestCalls + 1
    Err.Raise vbObjectError + 514, "test", "controlled plugin failure"
End Sub

Public Sub TestBatch()
    Dim sources As Collection, i As Long
    TestCalls = TestCalls + 1
    Set sources = StyledRanges(ActiveDocument)
    For i = sources.Count To 1 Step -1
        sources(i).Text = "converted"
    Next i
End Sub

Public Function RunRegressionTests() As String
    Dim doc As Document, f As Footnote, e As Endnote, r As Range
    Dim sources As Collection, okN As Long, failN As Long
    Set doc = ActiveDocument
    doc.Styles.Add Name:="MathSource", Type:=wdStyleTypeCharacter
    doc.Content.Text = "$main$"
    doc.Range(0, 6).Style = "MathSource"
    Set f = doc.Footnotes.Add(Range:=doc.Range(doc.Content.End - 1, doc.Content.End - 1), Text:="$foot$")
    f.Range.Style = "MathSource"
    Set e = doc.Endnotes.Add(Range:=doc.Range(doc.Content.End - 1, doc.Content.End - 1), Text:="$end$")
    e.Range.Style = "MathSource"
    Require CountPending(doc) = 3, "must count main, footnote, endnote"
    SelectLatexEquations
    Set sources = StyledRanges(doc)
    For Each r In sources
        Require r.HighlightColorIndex = wdYellow, "all stories must be highlighted"
    Next r
    ProbeEquationMacros
    Require TestCalls = 0, "probe must not execute callbacks"
    TestAnswer = vbNo
    ConvertWith "Md2DocxTools.TestConvertSource", "test"
    Require TestCalls = 0 And CountPending(doc) = 3, "cancel must preserve all sources"
    TestAnswer = vbYes
    ConvertWith "Md2DocxTools.TestConvertSource", "test"
    Require TestCalls = 3 And CountPending(doc) = 0, "must convert all stories once"
    doc.Content.Delete
    doc.Content.Text = "$a$ $b$"
    doc.Content.Style = doc.Styles(wdStyleDefaultParagraphFont)
    doc.Range(0, 3).Style = "MathSource"
    doc.Range(4, 7).Style = "MathSource"
    Require CountPending(doc) = 2, "batch fixture must have two sources"
    TestCalls = 0
    ConvertPending doc, "Md2DocxTools.TestBatch", okN, failN
    Require TestCalls = 1 And okN = 2 And CountPending(doc) = 0, "batch must skip already converted snapshots"
    doc.Content.Delete
    doc.Content.Text = "$retry$"
    doc.Range(0, 7).Style = "MathSource"
    TestCalls = 0
    okN = 0
    failN = 0
    ConvertPending doc, "Md2DocxTools.TestNoOp", okN, failN
    Require TestCalls = 1 And failN = 1 And okN = 0, "no-op must be bounded and counted as failed"
    TestCalls = 0
    okN = 0
    failN = 0
    ConvertPending doc, "Md2DocxTools.TestFailure", okN, failN
    Require TestCalls = 1 And failN = 1 And CountPending(doc) = 1, "failure must remain retryable"
    CleanupAfterConvert
    Require CountPending(doc) = 1 And InStr(doc.Content.Text, "$retry$") > 0, "cleanup must preserve pending delimiters"
    RunRegressionTests = "PASS: probe, confirmation, all stories, failures, cleanup"
End Function
'@
    $module.CodeModule.AddFromString($source + $tests)
    $result = $word.Run('Md2DocxTools.RunRegressionTests')
    if ($result -notlike 'PASS:*') { throw "Unexpected result: $result" }
    Write-Output $result
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    if ($null -ne $word) { $word.Quit() }
}
