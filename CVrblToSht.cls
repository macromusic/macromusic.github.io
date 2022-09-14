VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVrblToSht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfnclonIsArrayExtd
'Includes PfnclonArrayDimension
'Includes PfncobjInputBoxRange
'Includes PfixWriteArrayVariablesToWorkSheet

Option Explicit
Option Base 1

'◆ClassProc名_変数情報をエクセルシートに書き出す
'Rev.010

'//モジュール内定数
  Private Const meMstrMdlName As String = "CVrblToSht"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertyGet_出力データ
  Private myXobjPstdRng As Object
  
'//PropertySetLet_入力制御信号
  Private myXbisInptBxOFF As Boolean
  Private myXbisEachWrtON As Boolean
  
'//PropertySetLet_入力データ
  Private myZvarVrbl As Variant, myXobjPstFrstCell As Object
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myZvarPstData As Variant

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myZvarPstData = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getCmpltFlag() As Boolean
    getCmpltFlag = myXbisCmpltFlag
End Property

Public Property Get getPstdRng() As Object
    Set getPstdRng = myXobjPstdRng
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    Set myXobjPstdRng = Nothing
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
  Dim myXlonRows As Long, myXlonCols As Long
    myXlonRows = UBound(myZvarPstData, 1) - LBound(myZvarPstData, 1) + 1
    myXlonCols = UBound(myZvarPstData, 2) - LBound(myZvarPstData, 2) + 1
    Set myXobjPstdRng = myXobjPstFrstCell.Resize(myXlonRows, myXlonCols)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letVrbl(ByVal RHS As Variant)
    myZvarVrbl = RHS
End Property

Public Property Set setPstFrstCell(ByVal RHS As Object)
    Set myXobjPstFrstCell = RHS
End Property

Public Property Let letInptBxOFF(ByVal RHS As Boolean)
    myXbisInptBxOFF = RHS
'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
End Property

Public Property Let letEachWrtON(ByVal RHS As Boolean)
    myXbisEachWrtON = RHS
'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables1()
    myXbisExitFlag = False

'//入力変数が配列で空かどうかチェック
  Dim myXlonChck As Long: myXlonChck = PfnclonIsArrayExtd(myZvarVrbl)
    Select Case myXlonChck
        Case 0      '空の配列変数
            GoTo ExitPath
        Case 1      '空ではない配列変数
        Case -1     '配列ではない変数
    End Select
    
'//配列変数の次元数を取得
  Dim myXlonAryDmnsn As Long: myXlonAryDmnsn = PfnclonArrayDimension(myZvarVrbl)
    
'//配列変数の次元数に応じて処理を実施
  Dim i As Long, Li As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
    If myXlonAryDmnsn = 0 Then
        If CStr(myZvarVrbl) = "" Then GoTo ExitPath
        ReDim myZvarPstData(Lo, Lo) As Variant
        myZvarPstData(Lo, Lo) = myZvarVrbl
        
    ElseIf myXlonAryDmnsn = 1 Then
        Li = LBound(myZvarVrbl)
        i = UBound(myZvarVrbl) + Lo - Li
        ReDim myZvarPstData(i, Lo) As Variant
        For i = LBound(myZvarVrbl) To UBound(myZvarVrbl)
            myZvarPstData(i + Lo - Li, Lo) = myZvarVrbl(i)
        Next i
        
    ElseIf myXlonAryDmnsn = 2 Then
        myZvarPstData = myZvarVrbl
        
    Else
        GoTo ExitPath
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables2()
    myXbisExitFlag = False
    
    If Not myXobjPstFrstCell Is Nothing Then
        Set myXobjPstFrstCell = myXobjPstFrstCell.Cells(1, 1)
        Exit Sub
    End If
    If myXbisInptBxOFF = True Then GoTo ExitPath
    
  Dim myXstrPrompt As String, myXstrTitle As String, myXstrDefault As String
    myXstrPrompt = " " & vbCrLf _
                    & "データを貼り付ける位置を指定して下さい。" & vbCrLf _
                    & " "
    myXstrTitle = "貼り付け位置の指定"
    myXstrDefault = Selection.Address(ReferenceStyle:=Application.ReferenceStyle)
    Set myXobjPstFrstCell = PfncobjInputBoxRange( _
                            myXstrPrompt, myXstrTitle, myXstrDefault)
    If myXobjPstFrstCell Is Nothing Then GoTo ExitPath
    Set myXobjPstFrstCell = myXobjPstFrstCell.Cells(1, 1)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    Call ctrProc
End Sub

'PublicF_
Public Function fncbisCmpltFlag() As Boolean
    fncbisCmpltFlag = Empty
    Call ctrProc
    fncbisCmpltFlag = myXbisCmpltFlag
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    
    Call checkInputVariables1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
    Call checkInputVariables2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag

'//2次元配列変数のデータ一覧をシートに一括書出し
    Call PfixWriteArrayVariablesToWorkSheet(myXbisExitFlag, _
            myXobjPstFrstCell, myZvarPstData, myXbisEachWrtON)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"    'PassFlag

    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

 '定型Ｆ_指定変数が配列で空かどうかチェックする
Private Function PfnclonIsArrayExtd(ByRef myZvarOrgData As Variant) As Long
'PfnclonIsArrayExtd = 1  : 空ではない配列変数
'PfnclonIsArrayExtd = 0  : 空の配列変数
'PfnclonIsArrayExtd = -1 : 配列ではない変数
    PfnclonIsArrayExtd = Empty
  Dim myXlonChckFlag As Long, myXlonTmp As Long
'//配列であることを確認
    If IsArray(myZvarOrgData) = False Then
        myXlonChckFlag = -1: GoTo JumpPath
    End If
'//配列が空でないことを確認
    myXlonChckFlag = 0
    On Error GoTo JumpPath
    myXlonTmp = UBound(myZvarOrgData) - LBound(myZvarOrgData) + 1
    On Error GoTo 0
    If myXlonTmp <= 0 Then GoTo JumpPath
    myXlonChckFlag = 1
JumpPath:
    PfnclonIsArrayExtd = myXlonChckFlag
End Function

 '定型Ｆ_配列変数の次元数を取得する
Private Function PfnclonArrayDimension(ByRef myZvarOrgData As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarOrgData) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
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
'    Debug.Print "入力値: " & PfncobjInputBoxRange.Address( _
'                            ReferenceStyle:=Application.ReferenceStyle, _
'                            RowAbsolute:=True, ColumnAbsolute:=True, external:=True)
End Function

 '定型Ｐ_2次元配列変数のデータ一覧をシートに一括書出しする
Private Sub PfixWriteArrayVariablesToWorkSheet( _
            myXbisExitFlag As Boolean, _
            ByVal myXobjFrstCell As Object, ByRef myZvarOrgData As Variant, _
            Optional ByVal coXbisEachWrtON As Boolean = False)
    myXbisExitFlag = False
    If myXobjFrstCell Is Nothing Then GoTo ExitPath
  Dim Li As Long, myXvarTmp As Variant
    On Error GoTo ExitPath
    Li = LBound(myZvarOrgData, 1): myXvarTmp = myZvarOrgData(Li, Li)
    On Error GoTo 0
'//指定範囲内に配列変数が書出し可能かチェック
  Dim myXlonDataRows As Long, myXlonDataCols As Long
    myXlonDataRows = UBound(myZvarOrgData, 1) - LBound(myZvarOrgData, 1) + 1
    myXlonDataCols = UBound(myZvarOrgData, 2) - LBound(myZvarOrgData, 2) + 1
    If myXlonDataRows <= 0 Or myXlonDataCols <= 0 Then GoTo ExitPath
  Dim myXlonRngRows As Long, myXlonRngCols As Long
    With myXobjFrstCell
        myXlonRngRows = .Parent.Rows.Count - .Row + 1
        myXlonRngCols = .Parent.Columns.Count - .Column + 1
        If myXlonRngRows < myXlonDataRows Then GoTo ExitPath
        If myXlonRngCols < myXlonDataCols Then GoTo ExitPath
    End With
'//格納データをシートに書出し
    On Error GoTo ErrPath
    If coXbisEachWrtON = False Then
        myXobjFrstCell.Resize(myXlonDataRows, myXlonDataCols) = myZvarOrgData
    Else
      Dim i As Long, j As Long
        For i = LBound(myZvarOrgData, 1) To UBound(myZvarOrgData, 1)
            For j = LBound(myZvarOrgData, 2) To UBound(myZvarOrgData, 2)
                myXobjFrstCell.Offset(i - Li, j - Li).Value = myZvarOrgData(i, j)
            Next j
        Next i
    End If
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Exit Sub
ErrPath:
'//セルの書式を文字列に設定
    myXobjFrstCell.Resize(myXlonDataRows, myXlonDataCols).NumberFormatLocal = "@"
    Resume
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_変数情報をエクセルシートに書き出す
'Private Sub instCVrblToSht()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjPstdRng As Object
''//クラス内変数への入力
'  Dim myZvarVrbl As Variant, myXobjPstFrstCell As Object
'    myZvarVrbl = 1
'    Set myXobjPstFrstCell = Selection
'  Dim myXbisInptBxOFF As Boolean
'    myXbisInptBxOFF = False
'    'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
'    'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
'  Dim myXbisEachWrtON As Boolean
'    myXbisEachWrtON = False
'    'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
'    'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
''//CVrblToShtクラスモジュールのインスタンス生成
'  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
'    With myXinsVrblToSht
'    '//クラス内変数への入力
'        .letVrbl = myZvarVrbl
'        Set .setPstFrstCell = myXobjPstFrstCell
''        .letInptBxOFF = True
''        .letEachWrtON = True
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXbisCmpltFlag = .getCmpltFlag
'        Set myXobjPstdRng = .getPstdRng
'        myXbisCmpltFlag = .fncbisCmpltFlag
'    End With
'    Set myXinsVrblToSht = Nothing
'    myZvarVrbl = Empty: Set myXobjPstFrstCell = Nothing
'End Sub
