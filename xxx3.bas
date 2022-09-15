Attribute VB_Name = "xxx3"
Option Explicit
Option Base 1

 '定型Ｆ_指定ファイルのオブジェクトを取得する
Private Function PfncobjGetFile(ByVal myXstrFilePath As String) As Object
    Set PfncobjGetFile = Nothing
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FileExists(myXstrFilePath) = False Then Exit Function
        Set PfncobjGetFile = .GetFile(myXstrFilePath)
    End With
    Set myXobjFSO = Nothing
End Function

 '定型Ｆ_ファイルパスを指定してエクセルブックを開く
Private Function PfncobjOpenExcelBook( _
            ByVal myXstrFullName As String) As Object
    Set PfncobjOpenExcelBook = Nothing
    On Error Resume Next
    Set PfncobjOpenExcelBook = Workbooks.Open(myXstrFullName)
    On Error GoTo 0
End Function

 '抽象Ｐ_エクセルブック内の全シート＆全セル範囲＆全図形に対して処理を実行する
Private Sub PabsForEachSheetInBook( _
            myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object)
    myXbisExitFlag = False
  Dim myXlonShtCnt As Long: myXlonShtCnt = 0
  Dim myXobjSheet As Object
    For Each myXobjSheet In myXobjBook.Worksheets
    '//ブック内の全シートに対する処理
        Call PsubPreSheetOperation(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//シート内のデータ範囲に対する処理
        Call PsubForEachRangeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//シート内の全図形に対する処理
        Call PsubForEachShapeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//シート内の全グラフに対する処理
        Call PsubForEachChartInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//ブック内の全シートに対する処理
        myXlonShtCnt = myXlonShtCnt + 1
        Call PsubPostSheetOperation(myXobjSheet)
NextPath:
    Next
    Set myXobjSheet = Nothing
    myXbisExitFlag = False
    If myXlonShtCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubPreSheetOperation(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//ブック内の全シートに対する処理
'    XarbProgCode
End Sub
Private Sub PsubForEachRangeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//シート内のデータ範囲に対する処理
'//シート上のデータ範囲を取得
  Dim myXobjAllRng As Object
    With myXobjSheet
      Dim myXobjFrstRng As Object, myXobjLastRng As Object
        Set myXobjFrstRng = .Cells(1, 1)
        Set myXobjLastRng = .Cells.SpecialCells(xlCellTypeLastCell)
        Set myXobjAllRng = .Range(myXobjFrstRng, myXobjLastRng)
    End With
    Set myXobjFrstRng = Nothing: Set myXobjLastRng = Nothing
'//データ範囲を検索
  Dim myXlonRngCnt As Long: myXlonRngCnt = 0
  Dim myXobjRange As Object
    For Each myXobjRange In myXobjAllRng
        Call PsubRangeOperation(myXbisExitFlag, myXobjRange)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonRngCnt = myXlonRngCnt + 1
NextPath:
    Next
    Set myXobjAllRng = Nothing: Set myXobjRange = Nothing
    myXbisExitFlag = False
    If myXlonRngCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubForEachShapeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//シート内の全図形に対する処理
  Dim myXlonShpCnt As Long: myXlonShpCnt = 0
  Dim myXobjShape As Object
    For Each myXobjShape In myXobjSheet.Shapes
        Call PsubShapeOperation(myXbisExitFlag, myXobjShape)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonShpCnt = myXlonShpCnt + 1
NextPath:
    Next
    Set myXobjShape = Nothing
    myXbisExitFlag = False
    If myXlonShpCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubForEachChartInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//シート内の全グラフに対する処理
  Dim myXlonChrtCnt As Long: myXlonChrtCnt = 0
  Dim myXobjChrtObjct As Object
    For Each myXobjChrtObjct In myXobjSheet.Charts
        Call PsubChartOperation(myXbisExitFlag, myXobjChrtObjct)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonChrtCnt = myXlonChrtCnt + 1
NextPath:
    Next
    Set myXobjChrtObjct = Nothing
    myXbisExitFlag = False
    If myXlonChrtCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubRangeOperation(myXbisExitFlag As Boolean, myXobjRange As Object)
    myXbisExitFlag = False
'//シート内のデータ範囲に対する処理
'    XarbProgCode
End Sub
Private Sub PsubShapeOperation(myXbisExitFlag As Boolean, myXobjShape As Object)
    myXbisExitFlag = False
'//シート内の全図形に対する処理
'    XarbProgCode
End Sub
Private Sub PsubChartOperation(myXbisExitFlag As Boolean, myXobjChrtObjct As Object)
    myXbisExitFlag = False
'//シート内の全グラフに対する処理
'    XarbProgCode
End Sub
Private Sub PsubPostSheetOperation(myXobjSheet As Object)
'//ブック内の全シートに対する処理
'    XarbProgCode
End Sub

 '定型Ｐ_エクセルブックのプロパティのハイパーリンクの基点を取得する
Public Function PfncstrGetHyperLinkBase(ByVal myXobjBook As Object) As String
    PfncstrGetHyperLinkBase = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkBase = myXobjBook.BuiltinDocumentProperties("Hyperlink base").Value
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｆ_指定セル範囲に設定されたハイパーリンク先のパスを取得する
Public Function PfncstrGetHyperLinkPathAtRange(ByVal myXobjRange As Object) As String
    PfncstrGetHyperLinkPathAtRange = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkPathAtRange = myXobjRange.Hyperlinks(1).Address
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｆ_相対ファイルパスを指定して絶対パスを取得する
Public Function PfncstrGetAbsolutePath( _
            ByVal myXstrRltvPath As String, ByVal myXobjBook As Object, _
            Optional ByVal coXstrChar As String = "..") As String
    PfncstrGetAbsolutePath = Empty
    If myXstrRltvPath = "" Then Exit Function
    If InStr(myXstrRltvPath, "../") <> 0 Then _
        myXstrRltvPath = Replace(myXstrRltvPath, "../", "..\")
  Dim myXstrAbsltPath As String
  Dim myXstrPrntPath As String, myXstrChldPath As String
    myXstrPrntPath = myXobjBook.Path
    myXstrChldPath = myXstrRltvPath
  Dim i As Long, j As Long
  Dim m As Long, n As Long: m = 0: n = 0
    For i = 1 To Len(myXstrPrntPath)
        If Mid(myXstrPrntPath, i, Len("\")) = "\" Then m = m + 1
    Next i
    For j = 1 To Len(myXstrChldPath)
        If Mid(myXstrChldPath, i, Len(coXstrChar)) = ".." Then n = n + 1
    Next j
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If m >= n Then
        Do While Left(myXstrChldPath, Len("..")) = ".."
            myXstrPrntPath = myXobjFSO.GetParentFolderName(myXstrPrntPath)
            myXstrChldPath = Mid(myXstrChldPath, Len(coXstrChar) + 2)
        Loop
    End If
'    Debug.Print "親パス: " & myXstrPrntPath
'    Debug.Print "子パス: " & myXstrChldPath
    Select Case myXstrChldPath
        Case "": myXstrAbsltPath = myXstrPrntPath
        Case Else: myXstrAbsltPath = myXstrPrntPath & "\" & myXstrChldPath
    End Select
'    Debug.Print "絶対パス: " & myXstrAbsltPath
    PfncstrGetAbsolutePath = myXstrAbsltPath
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_エクセルブックのプロパティのハイパーリンクの基点を設定する
Private Sub PfixSetHyperLinkBase(myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object, ByVal myXstrHypLnkBase As String)
  Const coXstrBkPrptyName As String = "Hyperlink base"
    myXbisExitFlag = False
    If myXstrHypLnkBase = "" Then Exit Sub
    On Error GoTo ExitPath
    myXobjBook.BuiltinDocumentProperties(coXstrBkPrptyName).Value = myXstrHypLnkBase
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｐ_指定セル範囲にハイパーリンクを設定する
Private Sub PfixSetHyperLinkWithSheetCellAtRange(myXbisExitFlag As Boolean, _
            ByVal myXobjRange As Object, ByVal myXstrHypLnkAdrs As String, _
            ByVal myXstrSubAdrs As String, ByVal myXstrTxt As String)
'myXstrSubAdrs : "シート名!セル位置"
    myXbisExitFlag = False
    If myXobjRange Is Nothing Then Exit Sub
    If myXstrHypLnkAdrs = "" And myXstrSubAdrs = "" Then Exit Sub
    If myXstrTxt = "" Then myXstrTxt = myXobjRange.Value
    On Error GoTo ExitPath
    Call myXobjRange.Worksheet.Hyperlinks.Add( _
            Anchor:=myXobjRange, Address:=myXstrHypLnkAdrs, _
            SubAddress:=myXstrSubAdrs, TextToDisplay:=myXstrTxt)
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｆ_指定図形に設定されたハイパーリンク先のパスを取得する
Private Function PfncstrGetHyperLinkPathAtShape(ByVal myXobjShape As Object) As String
    PfncstrGetHyperLinkPathAtShape = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkPathAtShape = myXobjShape.Hyperlink.Address
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｐ_指定図形にハイパーリンクを設定する
Private Sub PfixSetHyperLinkAtShape(myXbisExitFlag As Boolean, _
            ByVal myXobjShape As Object, ByVal myXstrHypLnkAdrs As String)
    myXbisExitFlag = False
    If myXstrHypLnkAdrs = "" Then Exit Sub
    On Error GoTo ExitPath
    Call myXobjShape.Parent.Hyperlinks _
        .Add(Anchor:=myXobjShape, Address:=myXstrHypLnkAdrs)
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｐ_エクセルブックを上書き保存する
Private Sub PfixOverwriteSaveExcelBook(myXbisExitFlag As Boolean, _
            ByVal myXstrBookName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Workbooks(myXstrBookName).Save
    On Error GoTo 0
    GoTo ExitPath
ErrPath:
    myXbisExitFlag = True
ExitPath:
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
End Sub

 '定型Ｐ_エクセルブックを閉じる
Private Sub PfixCloseExcelBook(myXbisExitFlag As Boolean, _
            ByVal myXstrBookName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Workbooks(myXstrBookName).Close SaveChanges:=False
    On Error GoTo 0
    GoTo ExitPath
ErrPath:
    myXbisExitFlag = True
ExitPath:
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
End Sub
