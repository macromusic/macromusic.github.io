VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStrDcmtFltr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckArrayDimension
'Includes PfixVStringDecimationFilterRegularStep
'Includes PfixHStringDecimationFilterRegularStep

Option Explicit
Option Base 1

'◆ClassProc名_データ配列を間引き処理する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CStrDcmtFltr"
  
'//PropertyGet_出力データ
  Private myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
            myZstrDcmtData() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean
  
'//PropertySetLet_入力データ
  Private myZstrOrgDataIN() As String
  Private myXlonDcmtRegStep As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZstrOrgData() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getDcmtRowCnt() As Long
    getDcmtRowCnt = myXlonDcmtRowCnt
End Property

Public Property Get getDcmtColCnt() As Long
    getDcmtColCnt = myXlonDcmtColCnt
End Property

Public Property Get getDcmtData(ByVal i As Long, ByVal j As Long) As String
    getDcmtData = myZstrDcmtData(i, j)
'myZstrDcmtData(i, j) : 間引き後データ配列
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZstrDcmtData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letRowDrctn(ByVal RHS As Boolean)
    myXbisRowDrctn = RHS
'myXbisRowDrctn = True  : 行方向のデータを処理
'myXbisRowDrctn = False : 列方向のデータを処理
End Property

Public Property Let letOrgData(ByRef RHS() As String)
'myZstrOrgData(i, j) : 元データ配列
    myZstrOrgDataIN = RHS
End Property

Public Property Let letDcmtRegStep(ByVal RHS As Long)
    myXlonDcmtRegStep = RHS
End Property

Public Property Let letBgnRowCol(ByVal RHS As Long)
    myXlonBgnRowCol = RHS
End Property

Public Property Let letEndRowCol(ByVal RHS As Long)
    myXlonEndRowCol = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If PfncbisCheckArrayDimension(myZstrOrgDataIN, 2) = False Then GoTo ExitPath
    
  Dim L As Long, i As Long, j As Long
    L = LBound(myZstrOrgDataIN, 1)
    myXlonOrgRowCnt = UBound(myZstrOrgDataIN, 1) - L + 1
    myXlonOrgColCnt = UBound(myZstrOrgDataIN, 2) - L + 1
    ReDim myZstrOrgData(myXlonOrgRowCnt, myXlonOrgColCnt) As String
    For j = LBound(myZstrOrgDataIN, 2) To UBound(myZstrOrgDataIN, 2)
        For i = LBound(myZstrOrgDataIN, 1) To UBound(myZstrOrgDataIN, 1)
            myZstrOrgData(i + (1 - L), j + (1 - L)) = myZstrOrgDataIN(i, j)
        Next i
    Next j
    
    If myXlonDcmtRegStep <= 0 Then GoTo ExitPath
    
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
    
    If myXbisRowDrctn = True Then
    '//行方向に等間隔で間引く
        Call PfixVStringDecimationFilterRegularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZstrDcmtData, _
                myZstrOrgData, myXlonDcmtRegStep, myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
        
    Else
    '//列方向に等間隔で間引く
        Call PfixHStringDecimationFilterRegularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZstrDcmtData, _
                myZstrOrgData, myXlonDcmtRegStep, myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    End If
    
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

 '定型Ｆ_配列変数の次元数が指定次元と一致するかをチェックする
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarDataAry As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarDataAry) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarDataAry, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
End Function

 '定型Ｐ_2次元配列の行方向に等間隔で間引いたデータ配列を取得する
Private Sub PfixVStringDecimationFilterRegularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZstrDcmtData() As String, _
        ByRef myZstrOrgData() As String, ByVal myXlonDcmtStep As Long, _
        Optional ByVal coXlonBgnRow As Long = 1, _
        Optional ByVal coXlonEndRow As Long = 0)
'myZstrDcmtData(i, j) : 間引きデータ配列
'myZstrOrgData(i, j)  : 元データ配列
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZstrDcmtData
    On Error GoTo ExitPath
    If myXlonDcmtStep <= 0 Or myXlonDcmtStep > UBound(myZstrOrgData, 1) Then Exit Sub
    If coXlonBgnRow <= 0 Then coXlonBgnRow = LBound(myZstrOrgData, 1)
    If coXlonEndRow < coXlonBgnRow Or coXlonEndRow = UBound(myZstrOrgData, 1) Then _
        coXlonEndRow = UBound(myZstrOrgData, 1)
  Dim myXstrTmp As String: myXstrTmp = myZstrOrgData(coXlonBgnRow, 1)
    On Error GoTo 0
    myXlonDcmtRowCnt = Int((coXlonEndRow - coXlonBgnRow + 0.9) / myXlonDcmtStep) + 1
    myXlonDcmtColCnt = UBound(myZstrOrgData, 2)
    ReDim myZstrDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As String
  Dim i As Long, j As Long, n As Long: n = 0
    For i = coXlonBgnRow To coXlonEndRow Step myXlonDcmtStep
        n = n + 1
        For j = LBound(myZstrOrgData, 2) To UBound(myZstrOrgData, 2)
            myZstrDcmtData(n, j) = myZstrOrgData(i, j)
        Next j
    Next i
    Exit Sub
ExitPath:
End Sub

 '定型Ｐ_2次元配列の列方向に等間隔で間引いたデータ配列を取得する
Private Sub PfixHStringDecimationFilterRegularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZstrDcmtData() As String, _
        ByRef myZstrOrgData() As String, ByVal myXlonDcmtStep As Long, _
        Optional ByVal coXlonBgnCol As Long = 1, _
        Optional ByVal coXlonEndCol As Long = 0)
'myZstrDcmtData(i, j) : 間引きデータ配列
'myZstrOrgData(i, j)  : 元データ配列
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZstrDcmtData
    On Error GoTo ExitPath
    If myXlonDcmtStep <= 0 Or myXlonDcmtStep > UBound(myZstrOrgData, 2) Then Exit Sub
    If coXlonBgnCol <= 0 Then coXlonBgnCol = LBound(myZstrOrgData, 2)
    If coXlonEndCol < coXlonBgnCol Or coXlonEndCol = UBound(myZstrOrgData, 2) Then _
        coXlonEndCol = UBound(myZstrOrgData, 2)
  Dim myXstrTmp As String: myXstrTmp = myZstrOrgData(1, coXlonBgnCol)
    On Error GoTo 0
    myXlonDcmtRowCnt = UBound(myZstrOrgData, 1)
    myXlonDcmtColCnt = Int((coXlonEndCol - coXlonBgnCol + 0.9) / myXlonDcmtStep) + 1
    ReDim myZstrDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As String
  Dim i As Long, j As Long, n As Long: n = 0
    For j = coXlonBgnCol To coXlonEndCol Step myXlonDcmtStep
        n = n + 1
        For i = LBound(myZstrOrgData, 1) To UBound(myZstrOrgData, 1)
            myZstrDcmtData(i, n) = myZstrOrgData(i, j)
        Next i
    Next j
    Exit Sub
ExitPath:
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_データ配列を間引き処理する
'Private Sub instCStrDcmtFltr()
'  Const coXlonDcmtRegStep As Long = 10
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZstrOrgData() As String
'    myXlonOrgRowCnt = UBound(myZstrOrgData, 1)
'    myXlonOrgColCnt = UBound(myZstrOrgData, 2)
'  Dim myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, myZstrDcmtData() As String
''//CStrDcmtFltrクラスモジュールのインスタンス生成
'  Dim myXinsNumDcmtFltr As CStrDcmtFltr: Set myXinsNumDcmtFltr = New CStrDcmtFltr
'  Dim i As Long, j As Long
'    With myXinsNumDcmtFltr
'    '//クラス内変数への入力
'        .letRowDrctn = True
'        .letOrgData = myZdouOrgData
'        .letDcmtRegStep = coXlonDcmtRegStep
'        .letBgnRowCol = coXlonBgnRowCol
'        .letEndRowCol = coXlonEndRowCol
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonDcmtRowCnt = .getDcmtRowCnt
'        myXlonDcmtColCnt = .getDcmtColCnt
'        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo JumpPath
'        ReDim myZstrDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As String
'        For j = 1 To myXlonDcmtColCnt
'            For i = 1 To myXlonDcmtRowCnt
'                myZstrDcmtData(i, j) = .getDcmtData(i, j)
'            Next i
'        Next j
'    End With
'JumpPath:
'    Set myXinsNumDcmtFltr = Nothing
'    Call variablesOfCStrDcmtFltr(myXlonDcmtRowCnt, myZstrDcmtData)   'Debug.Print
'End Sub
'Private Sub variablesOfCStrDcmtFltr( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CStrDcmtFltrクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
