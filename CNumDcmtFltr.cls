VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNumDcmtFltr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes incPfixVNumericDecimationFilterIrregularStep
'Includes incPfixHNumericDecimationFilterIrregularStep
'Includes PfixQuotationRemainder
'Includes PfixVNumericDecimationFilterRegularStep
'Includes PfixHNumericDecimationFilterRegularStep

Option Explicit
Option Base 1

'◆ClassProc名_データ配列を間引き処理する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CNumDcmtFltr"
  
'//PropertyGet_出力データ
  Private myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
            myZdouDcmtData() As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisIrgStep As Boolean
  Private myXbisRowDrctn As Boolean
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonCalRowCol As Long, myXdouDcmtIrgStep As Double
  Private myXlonDcmtRegStep As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double

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

Public Property Get getDcmtData(ByVal i As Long, ByVal j As Long) As Double
    getDcmtData = myZdouDcmtData(i, j)
'myZdouDcmtData(i, j) : 間引き後データ配列
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZdouDcmtData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letIrgStep(ByVal RHS As Boolean)
    myXbisIrgStep = RHS
'myXbisIrgStep = True  : 指定間隔値に近い値で間引く
'myXbisIrgStep = False : 等間隔で間引く
End Property

Public Property Let letRowDrctn(ByVal RHS As Boolean)
    myXbisRowDrctn = RHS
'myXbisRowDrctn = True  : 行方向のデータを処理
'myXbisRowDrctn = False : 列方向のデータを処理
End Property

Public Property Let letOrgData(ByRef RHS() As Double)
'myZdouOrgDataIN(i, j) : 元データ配列
    myZdouOrgDataIN = RHS
End Property

Public Property Let letCalRowCol(ByVal RHS As Long)
    myXlonCalRowCol = RHS
End Property

Public Property Let letDcmtIrgStep(ByVal RHS As Double)
    myXdouDcmtIrgStep = RHS
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
    
    If PfncbisIsMatrix(myZdouOrgDataIN) = False Then GoTo ExitPath
    
  Dim L As Long, i As Long, j As Long
    L = LBound(myZdouOrgDataIN, 1)
    myXlonOrgRowCnt = UBound(myZdouOrgDataIN, 1) - L + 1
    myXlonOrgColCnt = UBound(myZdouOrgDataIN, 2) - L + 1
    ReDim myZdouOrgData(myXlonOrgRowCnt, myXlonOrgColCnt) As Double
    For j = LBound(myZdouOrgDataIN, 2) To UBound(myZdouOrgDataIN, 2)
        For i = LBound(myZdouOrgDataIN, 1) To UBound(myZdouOrgDataIN, 1)
            myZdouOrgData(i + (1 - L), j + (1 - L)) = myZdouOrgDataIN(i, j)
        Next i
    Next j
    
    If myXbisIrgStep = True Then
        If myXlonCalRowCol <= 0 Then GoTo ExitPath
        If myXdouDcmtIrgStep = 0 Then GoTo ExitPath
    Else
        If myXlonDcmtRegStep <= 0 Then GoTo ExitPath
    End If
    
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
        
    If myXbisIrgStep = True And myXbisRowDrctn = True Then
    '//指定列を指定間隔値に近い値で間引く
        Call incPfixVNumericDecimationFilterIrregularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZdouDcmtData, _
                myZdouOrgData, myXlonCalRowCol, myXdouDcmtIrgStep, _
                myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
        
    ElseIf myXbisIrgStep = True And myXbisRowDrctn = False Then
    '//指定行を指定間隔値に近い値で間引く
        Call incPfixHNumericDecimationFilterIrregularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZdouDcmtData, _
                myZdouOrgData, myXlonCalRowCol, myXdouDcmtIrgStep, _
                myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    ElseIf myXbisIrgStep = False And myXbisRowDrctn = True Then
    '//行方向に等間隔で間引く
        Call PfixVNumericDecimationFilterRegularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZdouDcmtData, _
                myZdouOrgData, myXlonDcmtRegStep, myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
        
    Else
    '//列方向に等間隔で間引く
        Call PfixHNumericDecimationFilterRegularStep( _
                myXlonDcmtRowCnt, myXlonDcmtColCnt, myZdouDcmtData, _
                myZdouOrgData, myXlonDcmtRegStep, myXlonBgnRowCol, myXlonEndRowCol)
        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    End If
    
ExitPath:
    Call initializeModuleVariables
End Sub

 '定型Ｆ_指定変数が数値行列かどうかをチェックする
Private Function PfncbisIsMatrix(ByRef myZvarMtrx As Variant) As Boolean
    PfncbisIsMatrix = False
  Dim myXdouTmp As Double, L1 As Long, L2 As Long
    On Error GoTo ExitPath
    L1 = LBound(myZvarMtrx, 1): L2 = LBound(myZvarMtrx, 2)
    myXdouTmp = myZvarMtrx(L1, L2)
    On Error GoTo 0
    PfncbisIsMatrix = True
ExitPath:
End Function

 '定型Ｐ_2次元の数値配列の指定列を指定間隔値に近い値で間引いたデータ配列を取得する
Private Sub incPfixVNumericDecimationFilterIrregularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZdouDcmtData() As Double, _
        ByRef myZdouOrgData() As Double, ByVal myXlonCalCol As Long, _
        ByVal myXdouDcmtStep As Double, _
        Optional ByVal coXlonBgnRow As Long = 1, _
        Optional ByVal coXlonEndRow As Long = 0)
'Includes PfixQuotationRemainder
'myZdouDcmtData(i, j) : 間引きデータ配列
'myZdouOrgData(i, j)  : 元データ配列
  Const coXlonTrgtCol As Long = 1
  Const coXlonDcmtCol As Long = 2
  Const coXlonTrgtVal As Long = 1
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZdouDcmtData
    On Error GoTo ExitPath
    If myXlonCalCol < LBound(myZdouOrgData, 2) _
        Or myXlonCalCol > UBound(myZdouOrgData, 2) Then Exit Sub
    If myXdouDcmtStep <= 0 Or myXdouDcmtStep > UBound(myZdouOrgData, 1) Then Exit Sub
    If coXlonBgnRow <= 0 Then coXlonBgnRow = LBound(myZdouOrgData, 1)
    If coXlonEndRow < coXlonBgnRow Or coXlonEndRow > UBound(myZdouOrgData, 1) Then _
        coXlonEndRow = UBound(myZdouOrgData, 1)
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(coXlonBgnRow, myXlonCalCol)
    On Error GoTo 0
'//指定間隔値に近い値にターゲットフラグを付ける
'myZdouCal(k, 1) : 間引き計算対象のデータ
'myZdouCal(k, 2) : ターゲットフラグ
  Dim myZdouCal() As Double, k As Long
    ReDim myZdouCal(coXlonEndRow - coXlonBgnRow + 1, 2) As Double
    For k = coXlonBgnRow To coXlonEndRow
        myZdouCal(k, coXlonDcmtCol) = myZdouOrgData(k, myXlonCalCol)
    Next k
  Dim myXlonTrgtCnt As Long
    Call PsubVNumericDecimation(myXlonTrgtCnt, myZdouCal, myXdouDcmtStep)
    If myXlonTrgtCnt <= 0 Then Exit Sub
'//ターゲットフラグの付いたデータを抜粋
    myXlonDcmtRowCnt = myXlonTrgtCnt
    myXlonDcmtColCnt = UBound(myZdouOrgData, 2)
    ReDim myZdouDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As Double
  Dim i As Long, j As Long, m As Long, n As Long: m = 0: n = 0
    For i = coXlonBgnRow To coXlonEndRow
        m = m + 1
        If myZdouCal(m, coXlonTrgtCol) <> coXlonTrgtVal Then GoTo NextPath
        n = n + 1
        For j = LBound(myZdouOrgData, 2) To UBound(myZdouOrgData, 2)
            myZdouDcmtData(n, j) = myZdouOrgData(i + coXlonBgnRow - 1, j)
        Next j
NextPath:
    Next i
    Erase myZdouCal
ExitPath:
End Sub
Private Sub PsubVNumericDecimation( _
            myXlonTrgtCnt As Long, myZdouCal() As Double, _
            ByVal myXlonStep As Double, _
            Optional ByVal coXlonTrgtCol As Long = 1, _
            Optional ByVal coXlonDcmtCol As Long = 2, _
            Optional ByVal coXlonTrgtVal As Long = 1)
'//指定間隔値に近い値にターゲットフラグを付ける
    myXlonTrgtCnt = Empty
    On Error GoTo ExitPath
  Dim m As Long: m = LBound(myZdouCal, 1)
  Dim n As Long: n = UBound(myZdouCal, 1)
  Dim p As Double: p = myZdouCal(m, coXlonDcmtCol)
  Dim q As Double: q = myZdouCal(n, coXlonDcmtCol)
    On Error GoTo 0
  Dim myXdouAveStep As Double, myXdouAveIntvl As Double, myXlonJump As Long
    myXdouAveStep = (q - p) / (n - m)
    myXdouAveIntvl = myXlonStep / myXdouAveStep
    myXlonJump = Int(myXdouAveIntvl) - 2
    If myXdouAveIntvl < 1 Then Exit Sub
  Dim k As Long: k = 0
  Dim myXdouQuot As Double, myXdouRemain As Double, myXdouNumer As Double
    myXdouNumer = myZdouCal(m, coXlonDcmtCol)
    Call PfixQuotationRemainder(myXdouQuot, myXdouRemain, myXdouNumer, myXlonStep)
    If myXdouRemain = 0 Then
        myZdouCal(m, coXlonTrgtCol) = coXlonTrgtVal
        k = k + 1
    End If
  Dim myXdouPrsntQ As Double, myXdouPrsntR As Double, _
        myXdouPrvsQ As Double, myXdouPrvsR As Double, i As Long
    For i = m + 1 To n
        myXdouNumer = myZdouCal(i, coXlonDcmtCol)
        Call PfixQuotationRemainder(myXdouPrsntQ, myXdouPrsntR, myXdouNumer, myXlonStep)
        If myXdouPrsntQ = myXdouQuot Then GoTo NextPath
        myXdouQuot = myXdouPrsntQ
        If myXdouPrsntR = 0 Then
            myZdouCal(i, coXlonTrgtCol) = coXlonTrgtVal
            k = k + 1: i = i + myXlonJump: GoTo NextPath
        End If
        myXdouNumer = myZdouCal(i - 1, coXlonDcmtCol)
        Call PfixQuotationRemainder(myXdouPrvsQ, myXdouPrvsR, myXdouNumer, myXlonStep)
        If myXdouPrsntR <= myXlonStep - myXdouPrvsR Then
            myZdouCal(i, coXlonTrgtCol) = coXlonTrgtVal
        Else
            myZdouCal(i - 1, coXlonTrgtCol) = coXlonTrgtVal
        End If
        k = k + 1: i = i + myXlonJump
NextPath:
    Next i
    myXlonTrgtCnt = k
ExitPath:
End Sub

 '定型Ｐ_2次元の数値配列の指定行を指定間隔値に近い値で間引いたデータ配列を取得する
Private Sub incPfixHNumericDecimationFilterIrregularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZdouDcmtData() As Double, _
        ByRef myZdouOrgData() As Double, ByVal myXlonCalRow As Long, _
        ByVal myXdouDcmtStep As Double, _
        Optional ByVal coXlonBgnCol As Long = 1, _
        Optional ByVal coXlonEndCol As Long = 0)
'Includes PfixQuotationRemainder
'myZdouDcmtData(i, j) : 間引きデータ配列
'myZdouOrgData(i, j)  : 元データ配列
  Const coXlonTrgtRow As Long = 1
  Const coXlonDcmtRow As Long = 2
  Const coXlonTrgtVal As Long = 1
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZdouDcmtData
    On Error GoTo ExitPath
    If myXlonCalRow < LBound(myZdouOrgData, 1) _
        Or myXlonCalRow > UBound(myZdouOrgData, 1) Then Exit Sub
    If myXdouDcmtStep <= 0 Or myXdouDcmtStep > UBound(myZdouOrgData, 2) Then Exit Sub
    If coXlonBgnCol <= 0 Then coXlonBgnCol = LBound(myZdouOrgData, 2)
    If coXlonEndCol < coXlonBgnCol Or coXlonEndCol = UBound(myZdouOrgData, 2) Then _
        coXlonEndCol = UBound(myZdouOrgData, 2)
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(myXlonCalRow, coXlonBgnCol)
    On Error GoTo 0
'//指定間隔値に近い値にターゲットフラグを付ける
'myZdouCal(1, k) : 間引き計算対象のデータ
'myZdouCal(2, k) : ターゲットフラグ
  Dim myZdouCal() As Double, k As Long
    ReDim myZdouCal(2, coXlonEndCol - coXlonBgnCol + 1) As Double
    For k = coXlonBgnCol To coXlonEndCol
        myZdouCal(coXlonDcmtRow, k) = myZdouOrgData(myXlonCalRow, k)
    Next k
  Dim myXlonTrgtCnt As Long
    Call PsubHNumericDecimation(myXlonTrgtCnt, myZdouCal, myXdouDcmtStep)
    If myXlonTrgtCnt <= 0 Then Exit Sub
'//ターゲットフラグの付いたデータを抜粋
    myXlonDcmtRowCnt = UBound(myZdouOrgData, 1)
    myXlonDcmtColCnt = myXlonTrgtCnt
    ReDim myZdouDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As Double
  Dim i As Long, j As Long, m As Long, n As Long: m = 0: n = 0
    For j = coXlonBgnCol To coXlonEndCol
        m = m + 1
        If myZdouCal(coXlonTrgtRow, m) <> coXlonTrgtVal Then GoTo NextPath
        n = n + 1
        For i = LBound(myZdouOrgData, 1) To UBound(myZdouOrgData, 1)
            myZdouDcmtData(i, n) = myZdouOrgData(i, j + coXlonBgnCol - 1)
        Next i
NextPath:
    Next j
    Erase myZdouCal
ExitPath:
End Sub
Private Sub PsubHNumericDecimation( _
            myXlonTrgtCnt As Long, myZdouCal() As Double, _
            ByVal myXlonStep As Double, _
            Optional ByVal coXlonTrgtRow As Long = 1, _
            Optional ByVal coXlonDcmtRow As Long = 2, _
            Optional ByVal coXlonTrgtVal As Long = 1)
'//指定間隔値に近い値にターゲットフラグを付ける
    myXlonTrgtCnt = Empty
    On Error GoTo ExitPath
  Dim m As Long: m = LBound(myZdouCal, 2)
  Dim n As Long: n = UBound(myZdouCal, 2)
  Dim p As Double: p = myZdouCal(coXlonDcmtRow, m)
  Dim q As Double: q = myZdouCal(coXlonDcmtRow, n)
    On Error GoTo 0
  Dim myXdouAveStep As Double, myXdouAveIntvl As Double, myXlonJump As Long
    myXdouAveStep = (q - p) / (n - m)
    myXdouAveIntvl = myXlonStep / myXdouAveStep
    myXlonJump = Int(myXdouAveIntvl) - 2
    If myXdouAveIntvl < 1 Then Exit Sub
  Dim k As Long: k = 0
  Dim myXdouQuot As Double, myXdouRemain As Double, myXdouNumer As Double
    myXdouNumer = myZdouCal(coXlonDcmtRow, m)
    Call PfixQuotationRemainder(myXdouQuot, myXdouRemain, myXdouNumer, myXlonStep)
    If myXdouRemain = 0 Then
        myZdouCal(coXlonTrgtRow, m) = coXlonTrgtVal
        k = k + 1
    End If
  Dim myXdouPrsntQ As Double, myXdouPrsntR As Double, _
        myXdouPrvsQ As Double, myXdouPrvsR As Double, j As Long
    For j = m + 1 To n
        myXdouNumer = myZdouCal(coXlonDcmtRow, j)
        Call PfixQuotationRemainder(myXdouPrsntQ, myXdouPrsntR, myXdouNumer, myXlonStep)
        If myXdouPrsntQ = myXdouQuot Then GoTo NextPath
        myXdouQuot = myXdouPrsntQ
        If myXdouPrsntR = 0 Then
            myZdouCal(coXlonTrgtRow, j) = coXlonTrgtVal
            k = k + 1: j = j + myXlonJump: GoTo NextPath
        End If
        myXdouNumer = myZdouCal(coXlonDcmtRow, j - 1)
        Call PfixQuotationRemainder(myXdouPrvsQ, myXdouPrvsR, myXdouNumer, myXlonStep)
        If myXdouPrsntR <= myXlonStep - myXdouPrvsR Then
            myZdouCal(coXlonTrgtRow, j) = coXlonTrgtVal
        Else
            myZdouCal(coXlonTrgtRow, j - 1) = coXlonTrgtVal
        End If
        k = k + 1: j = j + myXlonJump
NextPath:
    Next j
    myXlonTrgtCnt = k
ExitPath:
End Sub

 '定型Ｐ_商と余りを導出する
Private Sub PfixQuotationRemainder(myXdouQuot As Double, myXdouRemain As Double, _
        ByVal myXdouNumer As Double, ByVal myXdouDenom As Double)
'myXdouQuot   : 商
'myXdouRemain : 余り
'myXdouNumer  : 分子
'myXdouDenom  : 分母
    myXdouQuot = Empty: myXdouRemain = Empty
    If myXdouDenom = 0 Then Exit Sub
    On Error GoTo ExitPath
    myXdouQuot = Int(myXdouNumer / myXdouDenom)
    If myXdouQuot < 0 Then myXdouQuot = myXdouQuot + 1
    myXdouRemain = myXdouNumer - myXdouQuot * myXdouDenom
    On Error GoTo 0
ExitPath:
End Sub

 '定型Ｐ_2次元の数値配列の行方向に等間隔で間引いたデータ配列を取得する
Private Sub PfixVNumericDecimationFilterRegularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZdouDcmtData() As Double, _
        ByRef myZdouOrgData() As Double, ByVal myXlonDcmtStep As Long, _
        Optional ByVal coXlonBgnRow As Long = 1, _
        Optional ByVal coXlonEndRow As Long = 0)
'myZdouDcmtData(i, j) : 間引きデータ配列
'myZdouOrgData(i, j)  : 元データ配列
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZdouDcmtData
    On Error GoTo ExitPath
    If myXlonDcmtStep <= 0 Or myXlonDcmtStep > UBound(myZdouOrgData, 1) Then Exit Sub
    If coXlonBgnRow <= 0 Then coXlonBgnRow = LBound(myZdouOrgData, 1)
    If coXlonEndRow < coXlonBgnRow Or coXlonEndRow = UBound(myZdouOrgData, 1) Then _
        coXlonEndRow = UBound(myZdouOrgData, 1)
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(coXlonBgnRow, 1)
    On Error GoTo 0
    myXlonDcmtRowCnt = Int((coXlonEndRow - coXlonBgnRow + 0.9) / myXlonDcmtStep) + 1
    myXlonDcmtColCnt = UBound(myZdouOrgData, 2)
    ReDim myZdouDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As Double
  Dim i As Long, j As Long, n As Long: n = 0
    For i = coXlonBgnRow To coXlonEndRow Step myXlonDcmtStep
        n = n + 1
        For j = LBound(myZdouOrgData, 2) To UBound(myZdouOrgData, 2)
            myZdouDcmtData(n, j) = myZdouOrgData(i, j)
        Next j
    Next i
    Exit Sub
ExitPath:
End Sub

 '定型Ｐ_2次元の数値配列の列方向に等間隔で間引いたデータ配列を取得する
Private Sub PfixHNumericDecimationFilterRegularStep( _
        myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, _
        myZdouDcmtData() As Double, _
        ByRef myZdouOrgData() As Double, ByVal myXlonDcmtStep As Long, _
        Optional ByVal coXlonBgnCol As Long = 1, _
        Optional ByVal coXlonEndCol As Long = 0)
'myZdouDcmtData(i, j) : 間引きデータ配列
'myZdouOrgData(i, j)  : 元データ配列
    myXlonDcmtRowCnt = Empty: myXlonDcmtColCnt = Empty: Erase myZdouDcmtData
    On Error GoTo ExitPath
    If myXlonDcmtStep <= 0 Or myXlonDcmtStep > UBound(myZdouOrgData, 2) Then Exit Sub
    If coXlonBgnCol <= 0 Then coXlonBgnCol = LBound(myZdouOrgData, 2)
    If coXlonEndCol < coXlonBgnCol Or coXlonEndCol = UBound(myZdouOrgData, 2) Then _
        coXlonEndCol = UBound(myZdouOrgData, 2)
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(1, coXlonBgnCol)
    On Error GoTo 0
    myXlonDcmtRowCnt = UBound(myZdouOrgData, 1)
    myXlonDcmtColCnt = Int((coXlonEndCol - coXlonBgnCol + 0.9) / myXlonDcmtStep) + 1
    ReDim myZdouDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As Double
  Dim i As Long, j As Long, n As Long: n = 0
    For j = coXlonBgnCol To coXlonEndCol Step myXlonDcmtStep
        n = n + 1
        For i = LBound(myZdouOrgData, 1) To UBound(myZdouOrgData, 1)
            myZdouDcmtData(i, n) = myZdouOrgData(i, j)
        Next i
    Next j
    Exit Sub
ExitPath:
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub


''◆ClassProc名_データ配列を間引き処理する
'Private Sub instCNumDcmtFltr()
'  Const coXlonCalRowCol As Long = 2
'  Const coXdouDcmtIrgStep As Double = 0.2
'  Const coXlonDcmtRegStep As Long = 10
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
'    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
'    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
'  Dim myXlonDcmtRowCnt As Long, myXlonDcmtColCnt As Long, myZdouDcmtData() As Double
''//CNumDcmtFltrクラスモジュールのインスタンス生成
'  Dim myXinsNumDcmtFltr As CNumDcmtFltr: Set myXinsNumDcmtFltr = New CNumDcmtFltr
'  Dim i As Long, j As Long
'    With myXinsNumDcmtFltr
'    '//クラス内変数への入力
''        .letIrgStep = True
'        .letRowDrctn = True
'        .letOrgData = myZdouOrgData
'        .letCalRowCol = coXlonCalRowCol
'        .letDcmtIrgStep = coXdouDcmtIrgStep
'        .letDcmtRegStep = coXlonDcmtRegStep
'        .letBgnRowCol = coXlonBgnRowCol
'        .letEndRowCol = coXlonEndRowCol
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonDcmtRowCnt = .getDcmtRowCnt
'        myXlonDcmtColCnt = .getDcmtColCnt
'        If myXlonDcmtRowCnt <= 0 Or myXlonDcmtColCnt <= 0 Then GoTo JumpPath
'        ReDim myZdouDcmtData(myXlonDcmtRowCnt, myXlonDcmtColCnt) As Double
'        For j = 1 To myXlonDcmtColCnt
'            For i = 1 To myXlonDcmtRowCnt
'                myZdouDcmtData(i, j) = .getDcmtData(i, j)
'            Next i
'        Next j
'    End With
'JumpPath:
'    Set myXinsNumDcmtFltr = Nothing
'    Call variablesOfCNumDcmtFltr(myXlonDcmtRowCnt, myZdouDcmtData)   'Debug.Print
'End Sub
'Private Sub variablesOfCNumDcmtFltr( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CNumDcmtFltrクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
