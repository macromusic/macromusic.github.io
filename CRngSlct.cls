VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRngSlct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfixEnlargeSelection
'Includes PfncstrDefaultAddressInputBox
'Includes PfncobjInputBoxRange
'Includes PfixGetRangeBookSheetInformation
'Includes PfixGetRangeInformation

Option Explicit
Option Base 1

'◆ClassProc名_シート上のセル範囲を取得する
'Rev.008
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CRngSlct"
  
'//PropertyGet_出力データ
  Private myXobjSlctRng As Object
  Private myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
  Private myXlonBgnRow As Long, myXlonEndRow As Long, _
            myXlonBgnCol As Long, myXlonEndCol As Long, _
            myXlonRows As Long, myXlonCols As Long
  
'//PropertySetLet_入力制御信号
  Private myXlonRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getSlctRng() As Object
    Set getSlctRng = myXobjSlctRng
End Property

Public Property Get getBook() As Object
    Set getBook = myXobjBook
End Property

Public Property Get getShtName() As String
    getShtName = myXstrShtName
End Property

Public Property Get getShtNo() As Long
    getShtNo = myXlonShtNo
End Property

Public Property Get getBgnEndRowCol(ByVal i As Long, ByVal j As Long) As Long
    Select Case i & j
        Case "11": getBgnEndRowCol = myXlonBgnRow
        Case "21": getBgnEndRowCol = myXlonEndRow
        Case "12": getBgnEndRowCol = myXlonBgnCol
        Case "22": getBgnEndRowCol = myXlonEndCol
        Case "10": getBgnEndRowCol = myXlonRows
        Case "01": getBgnEndRowCol = myXlonCols
        Case Else
    End Select
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    Set myXobjSlctRng = Nothing
    Set myXobjBook = Nothing: myXstrShtName = Empty: myXlonShtNo = Empty
    myXlonBgnRow = Empty: myXlonEndRow = Empty
    myXlonBgnCol = Empty: myXlonEndCol = Empty
    myXlonRows = Empty: myXlonCols = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letRngOptn(ByVal RHS As Long)
    myXlonRngOptn = RHS
'myXlonRngOptn = 0  : 選択範囲
'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
'myXlonRngOptn = 3  : 全データ範囲
End Property

Public Property Let letInptBoxPrmptTtl(ByVal k As Long, ByVal RHS As String)
    Select Case k
        Case 1: myXstrInptBxPrmpt = RHS
        Case 2: myXstrInptBxTtl = RHS
        Case Else
    End Select
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXlonRngOptn < 0 Or myXlonRngOptn > 3 Then myXlonRngOptn = 0
    
    If myXstrInptBxPrmpt = "" Then _
        myXstrInptBxPrmpt = " " & vbCrLf & _
            "処理を実行したい範囲を指定して下さい。" & vbCrLf & " "
        
    If myXstrInptBxTtl = "" Then _
        myXstrInptBxTtl = "範囲の入力"
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    Call ctrProc
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//S:デフォルト範囲を設定
    
'//セル選択範囲を全データ範囲まで拡大
  Dim myXobjDfltRng As Object, myXobjFrstCell As Object, myXobjLastCell As Object
    Call PfixEnlargeSelection( _
            myXobjDfltRng, myXobjFrstCell, myXobjLastCell, _
            myXobjSlctRng, myXlonRngOptn)
    
'//InputBox用のデフォルト範囲を設定
  Dim myXstrDefault As String
    myXstrDefault = PfncstrDefaultAddressInputBox(myXobjDfltRng)
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:データ取得範囲を指定
  Dim myXbisExitLpFlag As Boolean: myXbisExitLpFlag = False
    Do While myXbisExitLpFlag = False
        Set myXobjSlctRng = PfncobjInputBoxRange(myXstrInptBxPrmpt, myXstrInptBxTtl, myXstrDefault)
        If myXobjSlctRng Is Nothing Then GoTo ExitPath
        If InStr(myXobjSlctRng.Address, ",") <= 0 Then
            myXbisExitLpFlag = True
        Else
            MsgBox "処理範囲は複数個所に分散させずに、連続範囲で指定して下さい。"
        End If
    Loop
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//S:データ取得範囲の情報を取得
    
'//指定セル範囲オブジェクトからブックとシートの情報を取得
    Call PfixGetRangeBookSheetInformation(myXobjBook, myXstrShtName, myXlonShtNo, myXobjSlctRng)
    
'//指定セル範囲オブジェクトから範囲情報を取得
    Call PfixGetRangeInformation( _
            myXlonBgnRow, myXlonEndRow, myXlonBgnCol, myXlonEndCol, myXlonRows, myXlonCols, _
            myXobjSlctRng)
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    Set myXobjDfltRng = Nothing: Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

 '定型Ｐ_セル選択範囲を全データ範囲まで拡大する
Private Sub PfixEnlargeSelection( _
            myXobjExtndRng As Object, _
            myXobjFrstCell As Object, myXobjLastCell As Object, _
            ByVal myXobjSlctRng As Object, _
            Optional ByVal coXlonRngOptn As Long = 0)
'coXlonRngOptn = 0  : 選択範囲
'coXlonRngOptn = 1  : 選択位置から最終行までの範囲
'coXlonRngOptn = 2  : 選択位置から最終列までの範囲
'coXlonRngOptn = 3  : 全データ範囲
    Set myXobjExtndRng = Nothing
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
    On Error GoTo ExitPath
    If myXobjSlctRng Is Nothing Then Set myXobjSlctRng = Selection
    On Error GoTo 0
  Dim myXlonRow1 As Long, myXlonRow2 As Long, myXlonCol1 As Long, myXlonCol2 As Long
    myXlonRow1 = myXobjSlctRng.Row
    myXlonRow2 = myXlonRow1 + myXobjSlctRng.Rows.Count - 1
    myXlonCol1 = myXobjSlctRng.Column
    myXlonCol2 = myXlonCol1 + myXobjSlctRng.Columns.Count - 1
  Dim myXlonEndRow As Long, myXlonEndCol As Long, myXlonTmp As Long
  Dim i As Long, j As Long
    With myXobjSlctRng
        If coXlonRngOptn = 1 Then
            With .Worksheet
                For j = myXlonCol1 To myXlonCol2
                    myXlonTmp = .Cells(.Rows.Count, j).End(xlUp).Row
                    If myXlonEndRow < myXlonRow1 Then myXlonEndRow = myXlonRow2
                    If myXlonEndRow < myXlonTmp Then myXlonEndRow = myXlonTmp
                Next j
                Set myXobjFrstCell = .Cells(myXlonRow1, myXlonCol1)
                Set myXobjLastCell = .Cells(myXlonEndRow, myXlonCol2)
            End With
       ElseIf coXlonRngOptn = 2 Then
            With .Worksheet
                For i = myXlonRow1 To myXlonRow2
                    myXlonTmp = .Cells(i, .Columns.Count).End(xlToLeft).Column
                    If myXlonEndCol < myXlonTmp Then myXlonEndCol = myXlonTmp
                    If myXlonEndCol < myXlonCol1 Then myXlonEndCol = myXlonCol2
                Next i
                Set myXobjFrstCell = .Cells(myXlonRow1, myXlonCol1)
                Set myXobjLastCell = .Cells(myXlonRow2, myXlonEndCol)
            End With
        ElseIf coXlonRngOptn = 3 Then
            Set myXobjFrstCell = .Worksheet.Cells(1, 1)
            Set myXobjLastCell = .Worksheet.Cells.SpecialCells(xlCellTypeLastCell)
        Else
            Set myXobjFrstCell = .Worksheet.Cells(myXlonRow1, myXlonCol1)
            Set myXobjLastCell = .Worksheet.Cells(myXlonRow2, myXlonCol2)
        End If
    Set myXobjExtndRng = .Worksheet.Range(myXobjFrstCell, myXobjLastCell)
    End With
ExitPath:
End Sub

 '定型Ｆ_InputBox用のデフォルト範囲を設定する
Private Function PfncstrDefaultAddressInputBox(ByVal myXobjDfltRng As Object) As String
    PfncstrDefaultAddressInputBox = Empty
  Dim myXstrDfltAdrs As String
    On Error GoTo ExitPath
    If myXobjDfltRng Is Nothing Then Set myXobjDfltRng = Selection
    myXstrDfltAdrs = myXobjDfltRng.Address(ReferenceStyle:=Application.ReferenceStyle, _
                        RowAbsolute:=True, ColumnAbsolute:=True, external:=False)
    On Error GoTo 0
    PfncstrDefaultAddressInputBox = myXstrDfltAdrs
ExitPath:
End Function

 '定型Ｆ_InputBoxメソッドでセル範囲の入力を促す
Private Function PfncobjInputBoxRange( _
                    ByVal myXstrPrompt As String, ByVal myXstrTitle As String, _
                    ByVal myXstrDefault As String) As Object
    Set PfncobjInputBoxRange = Nothing
    If myXstrTitle = "" Then myXstrTitle = "セル範囲の選択"
    If myXstrPrompt = "" Then myXstrPrompt = myXstrTitle
    On Error Resume Next
    Set PfncobjInputBoxRange = Application.InputBox( _
                                prompt:=myXstrPrompt, Title:=myXstrTitle, _
                                Default:=myXstrDefault, Type:=8)
    On Error GoTo 0
'    Debug.Print "入力値: " & PfncobjInputBoxRange.Address(ReferenceStyle:=Application.ReferenceStyle, _
'                                RowAbsolute:=True, ColumnAbsolute:=True, external:=True)    'Debug.Print
End Function

 '定型Ｐ_指定セル範囲オブジェクトからブックとシートの情報を取得する
Private Sub PfixGetRangeBookSheetInformation( _
            myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long, _
            ByVal myXobjOrgRng As Object)
    Set myXobjBook = Nothing: myXstrShtName = Empty: myXlonShtNo = Empty
    If myXobjOrgRng Is Nothing Then Exit Sub
    Set myXobjBook = myXobjOrgRng.Parent.Parent
    myXstrShtName = myXobjOrgRng.Parent.Name
  Dim i As Long
    For i = 1 To myXobjOrgRng.Parent.Parent.Worksheets.Count
        If myXobjOrgRng.Parent.Parent.Worksheets(i).Name = myXstrShtName Then
            myXlonShtNo = i: Exit For
        End If
    Next i
End Sub

 '定型Ｐ_指定セル範囲オブジェクトから範囲情報を取得する
Private Sub PfixGetRangeInformation( _
            myXlonBgnRow As Long, myXlonEndRow As Long, _
            myXlonBgnCol As Long, myXlonEndCol As Long, _
            myXlonRowCnt As Long, myXlonColCnt As Long, _
            ByVal myXobjOrgRng As Object)
    myXlonBgnRow = Empty: myXlonEndRow = Empty
    myXlonBgnCol = Empty: myXlonEndCol = Empty
    myXlonRowCnt = Empty: myXlonColCnt = Empty
    If myXobjOrgRng Is Nothing Then Exit Sub
    With myXobjOrgRng
        myXlonBgnRow = .Row
        myXlonRowCnt = .Rows.Count
        myXlonEndRow = .Row + myXlonRowCnt - 1
        myXlonBgnCol = .Column
        myXlonColCnt = .Columns.Count
        myXlonEndCol = .Column + myXlonColCnt - 1
    End With
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上のセル範囲を取得する
'Private Sub instCRngSlct()
'  Dim myXobjSlctRng As Object
'  Dim myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
'  Dim myXlonBgnRow As Long, myXlonEndRow As Long, _
'        myXlonBgnCol As Long, myXlonEndCol As Long, _
'        myXlonRows As Long, myXlonCols As Long
''//クラス内変数への入力
'  Dim myXlonRngOptn As Long: myXlonRngOptn = 0
'    'myXlonRngOptn = 0  : 選択範囲
'    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
'    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
'    'myXlonRngOptn = 3  : 全データ範囲
'  Dim myXstrInptBxPrmpt As String: myXstrInptBxPrmpt = ""
'  Dim myXstrInptBxTtl As String: myXstrInptBxTtl = ""
''//CRngSlctクラスモジュールのインスタンス生成
'  Dim myXinsRngSlct As CRngSlct: Set myXinsRngSlct = New CRngSlct
'    With myXinsRngSlct
'    '//クラス内変数への入力
'        .letRngOptn = myXlonRngOptn
'        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
'        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        Set myXobjSlctRng = .getSlctRng
'        If myXobjSlctRng Is Nothing Then GoTo JumpPath
'        Set myXobjBook = .getBook
'        myXstrShtName = .getShtName
'        myXlonShtNo = .getShtNo
'        myXlonBgnRow = .getBgnEndRowCol(1, 1)
'        myXlonEndRow = .getBgnEndRowCol(2, 1)
'        myXlonBgnCol = .getBgnEndRowCol(1, 2)
'        myXlonEndCol = .getBgnEndRowCol(2, 2)
'        myXlonRows = .getBgnEndRowCol(1, 0)
'        myXlonCols = .getBgnEndRowCol(0, 1)
'    End With
'JumpPath:
'    Set myXinsRngSlct = Nothing
'End Sub
