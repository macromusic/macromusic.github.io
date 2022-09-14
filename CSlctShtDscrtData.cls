VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSlctShtDscrtData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncstrDefaultAddressInputBox
'Includes PfncobjInputBoxRange
'Includes PfixGetRangeBookSheetInformation
'Includes PfixSplitRange

Option Explicit
Option Base 1

'◆ClassProc名_シート上の不連続範囲を指定してその範囲のデータと情報を取得する
'Rev.004
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSlctShtDscrtData"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonDataCnt As Long, myZobjDataCell() As Object, myZvarShtData() As Variant
    'myZvarShtData(i, 1) = Row
    'myZvarShtData(i, 2) = Column
    'myZvarShtData(i, 3) = SheetData
    'myZvarShtData(i, 4) = CommentData
  Private myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
  
'//PropertySetLet_入力制御信号
  Private myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean
  
'//PropertySetLet_入力データ
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjDataRng As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjDataRng = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getDataCnt() As Long
    getDataCnt = myXlonDataCnt
End Property

Public Property Get getDataCellAry(ByVal i As Long) As Object
    Set getDataCellAry = myZobjDataCell(i)
'myZobjDataCell(i) = セル配列
End Property

Public Property Get getShtDataAry(ByVal i As Long, ByVal j As Long) As Variant
    getShtDataAry = myZvarShtData(i, j)
'myZvarShtData(i, 1) = Row
'myZvarShtData(i, 2) = Column
'myZvarShtData(i, 3) = SheetData
'myZvarShtData(i, 4) = CommentData
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

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonDataCnt = Empty: Erase myZobjDataCell: Erase myZvarShtData
    Set myXobjBook = Nothing: myXstrShtName = Empty: myXlonShtNo = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letByVrnt(ByVal RHS As Boolean)
    myXbisByVrnt = RHS
'myXbisByVrnt = False : シートデータをStringで取得する
'myXbisByVrnt = True  : シートデータをVariantで取得する
End Property

Public Property Let letGetCmnt(ByVal RHS As Boolean)
    myXbisGetCmnt = RHS
'myXbisGetCmnt = False : コメントを取得しない
'myXbisGetCmnt = True  : コメントを取得する
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
    
'//S:データ取得範囲を設定
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:指定範囲からブック＆シートの情報を取得
    Call prsProcGetInformation
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:指定範囲のデータとコメントを取得
    Call prsProcGetData
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'SnsP_データ取得範囲を設定
Private Sub snsProc()
    myXbisExitFlag = False
    
'//InputBox用のデフォルト範囲を設定
  Dim myXobjDfltRng As Object, myXstrDefault As String
    Set myXobjDfltRng = Selection
    myXstrDefault = PfncstrDefaultAddressInputBox(myXobjDfltRng)

'//InputBoxメソッドでセル範囲の入力
    Set myXobjDataRng = PfncobjInputBoxRange(myXstrInptBxPrmpt, myXstrInptBxTtl, myXstrDefault)
    If myXobjDataRng Is Nothing Then GoTo ExitPath
    
    Set myXobjDfltRng = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_指定範囲からブック＆シートの情報を取得
Private Sub prsProcGetInformation()
    myXbisExitFlag = False
    
'//指定セル範囲オブジェクトからブックとシートの情報を取得
    Call PfixGetRangeBookSheetInformation(myXobjBook, myXstrShtName, myXlonShtNo, myXobjDataRng)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_指定範囲のデータとコメントを取得
Private Sub prsProcGetData()
    myXbisExitFlag = False
    
'//指定セル範囲を分割して各セルの配列を抽出
    Call PfixSplitRange(myXlonDataCnt, myZobjDataCell, myXobjDataRng)
    If myXlonDataCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4-1"   'PassFlag
    
'//出力配列に取得データを格納
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim i As Long: i = myXlonDataCnt + L - 1
    ReDim myZvarShtData(i, L + 3) As Variant
    
'//セルの配置を格納
    For i = LBound(myZobjDataCell) To UBound(myZobjDataCell)
        myZvarShtData(i, L + 0) = myZobjDataCell(i).Row
        myZvarShtData(i, L + 1) = myZobjDataCell(i).Column
    Next i

    If myXbisByVrnt = False Then
    '//セルの値をString型で格納
        For i = LBound(myZobjDataCell) To UBound(myZobjDataCell)
            myZvarShtData(i, L + 2) = CStr(myZobjDataCell(i).Value)
        Next i
        
    Else
    '//セルの値をVariant型で格納
        For i = LBound(myZobjDataCell) To UBound(myZobjDataCell)
            myZvarShtData(i, L + 2) = CVar(myZobjDataCell(i).Value)
        Next i
        
    End If
'    Debug.Print "PassFlag: " & meMstrMdlName & "4-2"   'PassFlag

  Dim myXstrTmpCmnt As String
    If myXbisGetCmnt = True Then
    '//セル配列から各セルのコメント内容を取得
        For i = LBound(myZobjDataCell) To UBound(myZobjDataCell)
            If TypeName(myZobjDataCell(i).Comment) = "Comment" Then
                myXstrTmpCmnt = myZobjDataCell(i).Comment.Text
            Else
                myXstrTmpCmnt = ""
            End If
            myZvarShtData(i, L + 3) = myXstrTmpCmnt
        Next i
    End If
'    Debug.Print "PassFlag: " & meMstrMdlName & "4-3"   'PassFlag
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

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

 '定型Ｐ_指定セル範囲を分割して各セルの配列を抽出する
Private Sub PfixSplitRange( _
            myXlonSpltCnt As Long, myZobjSpltRng() As Object, _
            ByVal myXobjOrgRng As Object)
'myZobjSpltRng(i) : 分割したセル配列
    myXlonSpltCnt = Empty: Erase myZobjSpltRng
    If myXobjOrgRng Is Nothing Then Exit Sub
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXvarTmpCell As Variant, n As Long: n = L - 1
    On Error Resume Next
    For Each myXvarTmpCell In myXobjOrgRng
        n = n + 1: ReDim Preserve myZobjSpltRng(n) As Object
        Set myZobjSpltRng(n) = myXvarTmpCell
    Next myXvarTmpCell
    On Error GoTo 0
    myXlonSpltCnt = n - L + 1
    myXvarTmpCell = Empty
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上の不連続範囲を指定してその範囲のデータと情報を取得する
'Private Sub instCSlctShtDscrtData()
'  Dim Lc As Long
'  Dim myXlonDataCnt As Long, myZobjDataCell() As Object, myZvarShtData() As Variant
'    'myZvarShtData(i, 1) = Row
'    'myZvarShtData(i, 2) = Column
'    'myZvarShtData(i, 3) = SheetData
'    'myZvarShtData(i, 4) = CommentData
'  Dim myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
''//クラス内変数への入力
'  Dim myXbisByVrnt As Boolean: myXbisByVrnt = False
'    'myXbisByVrnt = False : シートデータをStringで取得する
'    'myXbisByVrnt = True  : シートデータをVariantで取得する
'  Dim myXbisGetCmnt As Boolean: myXbisGetCmnt = True
'    'myXbisGetCmnt = False : コメントを取得しない
'    'myXbisGetCmnt = True  : コメントを取得する
'  Dim myXstrInptBxPrmpt As String: myXstrInptBxPrmpt = ""
'  Dim myXstrInptBxTtl As String: myXstrInptBxTtl = ""
''//CSlctShtDscrtDataクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long
'  Dim myXinsSlctShtDscrtData As CSlctShtDscrtData
'    Set myXinsSlctShtDscrtData = New CSlctShtDscrtData
'    With myXinsSlctShtDscrtData
'    '//クラス内変数への入力
'        .letByVrnt = myXbisByVrnt
'        .letGetCmnt = myXbisGetCmnt
'        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
'        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonDataCnt = .getDataCnt
'        If myXlonDataCnt <= 0 Then GoTo JumpPath
'        i = myXlonDataCnt + Lo - 1
'        ReDim myZobjDataCell(i) As Object
'        ReDim myZvarShtData(i, Lo + 3) As Variant
'        Lc = .getOptnBase
'        For i = 1 To myXlonDataCnt
'            Set myZobjDataCell(i + Lo - 1) = .getDataCellAry(i + Lc - 1)
'        Next i
'        For i = 1 To myXlonDataCnt
'            myZvarShtData(i + Lo - 1, Lo + 0) = .getShtDataAry(i + Lc - 1, Lc + 0)
'            myZvarShtData(i + Lo - 1, Lo + 1) = .getShtDataAry(i + Lc - 1, Lc + 1)
'            myZvarShtData(i + Lo - 1, Lo + 2) = .getShtDataAry(i + Lc - 1, Lc + 2)
'        Next i
'        If myXbisGetCmnt = True Then
'            For i = 1 To myXlonDataCnt
'                myZvarShtData(i + Lo - 1, Lo + 3) = .getShtDataAry(i + Lc - 1, Lo + 3)
'            Next i
'        End If
'        Set myXobjBook = .getBook
'        myXstrShtName = .getShtName
'        myXlonShtNo = .getShtNo
'    End With
'JumpPath:
'    Set myXinsSlctShtDscrtData = Nothing
'    Call variablesOfCSlctShtDscrtData(myXlonDataCnt, myZvarShtData)     'Debug.Print
'End Sub
'Private Sub variablesOfCSlctShtDscrtData( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSlctShtDscrtDataクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
