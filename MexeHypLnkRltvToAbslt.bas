Attribute VB_Name = "MexeHypLnkRltvToAbslt"
'Includes PincGetExcelBookObject
'Includes PfncstrGetHyperLinkBase
'Includes PfixSetHyperLinkBase
'Includes PabsForEachSheetInBook
'Includes PfncstrGetHyperLinkPathAtRange
'Includes PfncstrGetAbsolutePath
'Includes PfncbisCheckFolderExist
'Includes PfncbisCheckFileExist
'Includes PfixSetHyperLinkWithSheetCellAtRange
'Includes PfncstrGetHyperLinkPathAtShape
'Includes PfixSetHyperLinkAtShape
'Includes PfixOverwriteSaveExcelBook
'Includes PfixCloseExcelBook

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルブック内の全ハイパーリンクを絶対参照に変更する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "MexeHypLnkRltvToAbslt"

'//モジュール内定数_列挙体
Private Enum EnumX
'列挙体使用時の表記 : EnumX.rowX
'■myEnumの表記ルール
    '①シートNo. : "sht" & "Enum名" & " = " & "値" & "'シート名"
    '②行No.     : "row" & "Enum名" & " = " & "値" & "'検索するシート上の文字列"
    '③列No.     : "col" & "Enum名" & " = " & "値" & "'検索するシート上の文字列"
    '④行No.     : "row" & "Enum名" & " = " & "値" & "'comment" & "'検索するコメントの文字列"
    '⑤列No.     : "col" & "Enum名" & " = " & "値" & "'comment" & "'検索するコメントの文字列"
    shtX = 1        'Sheet1
    rowX = 1        '行No
    colX = 1        '列No
End Enum
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjBook As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjBook = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    
'//処理実行
    Call callMexeHypLnkToAbslt
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
'//処理実行
    Call ctrProc
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
End Sub

'CtrlP_
Private Sub ctrProc()
    myXbisCmpltFlag = False
    Call initializeModuleVariables
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
  Dim myXstrFullName As String
    myXstrFullName = "C:\Users\Hiroki\Documents\_VBA4XPC\11_プログラムデータベース\01_VBA構文\c10_ハイパーリンク" _
        & "\" & "test.xlsm"
'    myXstrFullName = ThisWorkbook.Worksheets(EnumX.shtX).Cells(EnumX.rowX, EnumX.colX).Value
    
'//指定エクセルブックの状態を確認してオブジェクトを取得
    Call PincGetExcelBookObject(myXbisExitFlag, myXobjBook, myXstrFullName)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "データ: " & myXobjBook.Name

'//エクセルブックのプロパティのハイパーリンクの基点を取得
  Dim myXstrHypLnkBase As String
    myXstrHypLnkBase = PfncstrGetHyperLinkBase(myXobjBook)
    If myXstrHypLnkBase <> "" Then GoTo ExitPath

'//エクセルブックのプロパティのハイパーリンクの基点を設定
    myXstrHypLnkBase = "*"
    Call PfixSetHyperLinkBase(myXbisExitFlag, myXobjBook, myXstrHypLnkBase)
    If myXbisExitFlag = True Then GoTo ExitPath

'//エクセルブック内の任意の同一全オブジェクトに対して処理を実行
    Call PabsForEachSheetInBook(myXbisExitFlag, myXobjBook)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag

'//エクセルブックを上書き保存
  Dim myXstrBookName As String
    myXstrBookName = myXobjBook.Name
    Call PfixOverwriteSaveExcelBook(myXbisExitFlag, myXstrBookName)
    If myXbisExitFlag = True Then GoTo ExitPath

'//エクセルブックを閉じる
    Call PfixCloseExcelBook(myXbisExitFlag, myXstrBookName)
    If myXbisExitFlag = True Then GoTo ExitPath
    
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

 '抽象Ｐ_エクセルブック内の全シート＆全セル範囲＆全図形に対して処理を実行する
Private Sub PabsForEachSheetInBook( _
            myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object)
    myXbisExitFlag = False
  Dim myXlonShtCnt As Long: myXlonShtCnt = 0
  Dim myXobjSheet As Object
    For Each myXobjSheet In myXobjBook.Worksheets
        myXlonShtCnt = myXlonShtCnt + 1
    '//シート内のデータ範囲に対する処理
        Call PsubForEachRangeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
'    '//シート内の全図形に対する処理
'        Call PsubForEachShapeInSheet(myXbisExitFlag, myXobjSheet)
'        If myXbisExitFlag = True Then GoTo NextPath
NextPath:
    Next
    Set myXobjSheet = Nothing
    myXbisExitFlag = False
    If myXlonShtCnt = 0 Then myXbisExitFlag = True
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
'Private Sub PsubForEachShapeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
'    myXbisExitFlag = False
''//シート内の全図形に対する処理
'  Dim myXlonShpCnt As Long: myXlonShpCnt = 0
'  Dim myXobjShape As Object
'    For Each myXobjShape In myXobjSheet.Shapes
'        Call PsubShapeOperation(myXbisExitFlag, myXobjShape)
'        If myXbisExitFlag = True Then GoTo NextPath
'        myXlonShpCnt = myXlonShpCnt + 1
'NextPath:
'    Next
'    Set myXobjShape = Nothing
'    myXbisExitFlag = False
'    If myXlonShpCnt = 0 Then myXbisExitFlag = True
'End Sub
Private Sub PsubRangeOperation(myXbisExitFlag As Boolean, myXobjRange As Object)
    myXbisExitFlag = False
'//シート内のデータ範囲に対する処理
    
'//指定セル範囲に設定されたハイパーリンク先のパスを取得
  Dim myXstrLinkPath As String
    myXstrLinkPath = PfncstrGetHyperLinkPathAtRange(myXobjRange)
    If myXstrLinkPath = "" Then Exit Sub
    
'//相対ファイルパスを指定して絶対パスを取得
  Dim myXstrRltvPath As String
    myXstrRltvPath = myXstrLinkPath
  Dim myXstrAbsltPath As String
    myXstrAbsltPath = PfncstrGetAbsolutePath(myXstrRltvPath, myXobjBook)

'//指定フォルダの存在を確認
  Dim myXbisFldrExistFlag As Boolean
    myXbisFldrExistFlag = PfncbisCheckFolderExist(myXstrAbsltPath)

'//指定ファイルの存在を確認
  Dim myXbisFileExistFlag As Boolean
    myXbisFileExistFlag = PfncbisCheckFileExist(myXstrAbsltPath)
    
    If myXbisFldrExistFlag = False And myXbisFldrExistFlag = False Then
        Debug.Print "パスエラー: " & myXstrRltvPath
        myXbisExitFlag = True
        Exit Sub
    End If

'//指定セル範囲にハイパーリンクを設定
  Dim myXstrHypLnkAdrs As String, myXstrSubAdrs As String, myXstrTxt As String
    myXstrHypLnkAdrs = myXstrAbsltPath
    myXstrSubAdrs = ""
    myXstrTxt = ""
    Call PfixSetHyperLinkWithSheetCellAtRange( _
            myXbisExitFlag, _
            myXobjRange, myXstrHypLnkAdrs, myXstrSubAdrs, myXstrTxt)
    If myXbisExitFlag = True Then Exit Sub

End Sub
'Private Sub PsubShapeOperation(myXbisExitFlag As Boolean, myXobjShape As Object)
'    myXbisExitFlag = False
''//シート内の全図形に対する処理
''    XarbProgCode
'End Sub
'End Sub

'===============================================================================================

 '定型Ｐ_指定エクセルブックの状態を確認してオブジェクトを取得する
Private Sub PincGetExcelBookObject( _
            myXbisExitFlag As Boolean, myXobjBook As Object, _
            ByVal myXstrFullName As String)
'Includes PfnclonCheckExcelBookOpening
'Includes PfncobjOpenExcelBook
'Includes PfncobjGetFile
'Includes PfixCloseExcelBook
'Includes PfncobjGetExcelBookIfOpened
    myXbisExitFlag = False: Set myXobjBook = Nothing
  Dim myXlonCheckBookOpening As Long, myXstrBookName As String
'//指定エクセルブックが既に開いているか確認
    myXlonCheckBookOpening = PfnclonCheckExcelBookOpening(myXstrFullName)
    Select Case myXlonCheckBookOpening
        Case 0
            myXbisExitFlag = True
            Exit Sub
        Case 1
        '//ファイルパスを指定してエクセルブックを開く
            Set myXobjBook = PfncobjOpenExcelBook(myXstrFullName)
        Case 2
        '//指定ファイルのオブジェクトを取得
            Set myXobjBook = PfncobjGetFile(myXstrFullName)
            myXstrBookName = myXobjBook.Name
        '//エクセルブックを閉じる
            Call PfixCloseExcelBook(myXbisExitFlag, myXstrBookName)
            If myXbisExitFlag = True Then Exit Sub
        '//ファイルパスを指定してエクセルブックを開く
            Set myXobjBook = PfncobjOpenExcelBook(myXstrFullName)
        Case 3
        '//指定ファイルのオブジェクトを取得
            Set myXobjBook = PfncobjGetFile(myXstrFullName)
            myXstrBookName = myXobjBook.Name
        '//指定名のエクセルブックが既に開いていればブックオブジェクトを取得
            Set myXobjBook = PfncobjGetExcelBookIfOpened(myXstrBookName)
        Case Else
            Exit Sub
    End Select
End Sub

 '定型Ｆ_指定エクセルブックが既に開いているか確認する
Private Function PfnclonCheckExcelBookOpening( _
            ByVal myXstrFullName As String) As Long
'PfnclonCheckExcelBookOpening = 0 : 指定ブックが存在しない
'PfnclonCheckExcelBookOpening = 1 : 開いていない
'PfnclonCheckExcelBookOpening = 2 : 指定ブックと同一名の別ブックが開いている
'PfnclonCheckExcelBookOpening = 3 : 指定ブックが開いている
    PfnclonCheckExcelBookOpening = Empty
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If myXobjFSO.FileExists(myXstrFullName) = False Then Exit Function
  Dim myXstrBookName As String
    myXstrBookName = myXobjFSO.GetFileName(myXstrFullName)
    On Error GoTo ExitPath
  Dim myXstrTmp As String: myXstrTmp = Workbooks(myXstrBookName).FullName
    On Error GoTo 0
    If myXstrTmp = myXstrFullName Then
        PfnclonCheckExcelBookOpening = 3
    Else
        PfnclonCheckExcelBookOpening = 2
    End If
    Set myXobjFSO = Nothing
    Exit Function
ExitPath:
    PfnclonCheckExcelBookOpening = 1
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

 '定型Ｆ_指定名のエクセルブックが既に開いていればブックオブジェクトを取得する
Private Function PfncobjGetExcelBookIfOpened( _
            ByVal myXstrBookName As String) As Object
    Set PfncobjGetExcelBookIfOpened = Nothing
    On Error GoTo ExitPath
    Set PfncobjGetExcelBookIfOpened = Workbooks(myXstrBookName)
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｐ_エクセルブックのプロパティのハイパーリンクの基点を取得する
Private Function PfncstrGetHyperLinkBase(ByVal myXobjBook As Object) As String
    PfncstrGetHyperLinkBase = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkBase = myXobjBook.BuiltinDocumentProperties("Hyperlink base").Value
    On Error GoTo 0
ExitPath:
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

 '定型Ｆ_指定セル範囲に設定されたハイパーリンク先のパスを取得する
Private Function PfncstrGetHyperLinkPathAtRange(ByVal myXobjRange As Object) As String
    PfncstrGetHyperLinkPathAtRange = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkPathAtRange = myXobjRange.Hyperlinks(1).Address
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｆ_相対ファイルパスを指定して絶対パスを取得する
Private Function PfncstrGetAbsolutePath( _
            ByVal myXstrRltvPath As String, ByVal myXobjBook As Object) As String
    PfncstrGetAbsolutePath = Empty
    If myXstrRltvPath = "" Then Exit Function
    If myXobjBook Is Nothing Then Exit Function
  Dim myXstrAbsltPath As String
  Dim myXstrPrntPath As String, myXstrChldPath As String
    myXstrPrntPath = myXobjBook.Path
    myXstrChldPath = myXstrRltvPath
    myXstrPrntPath = Replace(myXstrPrntPath, "/", "\")
    myXstrChldPath = Replace(myXstrChldPath, "/", "\")
  Dim i As Long, j As Long, m As Long, n As Long: m = 0: n = 0
    If Left(myXstrChldPath, Len("..")) = ".." Then
        For i = 1 To Len(myXstrPrntPath)
            If Mid(myXstrPrntPath, i, Len("\")) = "\" Then m = m + 1
        Next i
        For j = 1 To Len(myXstrChldPath)
            If Mid(myXstrChldPath, i, Len("..")) = ".." Then n = n + 1
        Next j
        If m >= n Then
          Dim myXobjFSO As Object
            Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
            Do While Left(myXstrChldPath, Len("..")) = ".."
                myXstrPrntPath = myXobjFSO.GetParentFolderName(myXstrPrntPath)
                myXstrChldPath = Mid(myXstrChldPath, Len("..") + 2)
            Loop
            Set myXobjFSO = Nothing
        Else
            Exit Function
        End If
    End If
    Select Case myXstrChldPath
        Case "": myXstrAbsltPath = myXstrPrntPath
        Case Else: myXstrAbsltPath = myXstrPrntPath & "\" & myXstrChldPath
    End Select
'    Debug.Print "親パス: " & myXstrPrntPath
'    Debug.Print "子パス: " & myXstrChldPath
'    Debug.Print "絶対パス: " & myXstrAbsltPath
    PfncstrGetAbsolutePath = myXstrAbsltPath
End Function

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function
 
 '定型Ｆ_指定ファイルの存在を確認する
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

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

'DummyＰ_
Private Sub MsubDummy()
End Sub

'===============================================================================================

'◆ModuleProc名_エクセルブック内の全シート＆全セル範囲＆全図形に対して処理を実行する
Private Sub callMexeHypLnkToAbslt()
  Dim myXbisCompFlag As Boolean
    Call MexeHypLnkToAbslt.callProc(myXbisCompFlag)
    Debug.Print "完了: " & myXbisCompFlag
End Sub
