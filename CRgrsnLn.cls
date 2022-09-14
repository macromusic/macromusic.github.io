VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRgrsnLn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes PfixArrayPartRowNumeric
'Includes PfixArrayPartColumnNumeric
'Includes incPfixRegressionLineYArray
'Includes incPfixRegressionLineXArray
'Includes PfixSlopeIntercept

Option Explicit
Option Base 1

'◆ClassProc名_データ配列の全区間で回帰直線処理をしてデータを取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CRgrsnLn"
  
'//PropertyGet_出力データ
  Private myXlonStrtRowCol As Long, myXlonLastRowCol As Long
  Private myXlonCalRowCnt As Long, myXlonCalColCnt As Long, _
            myZdouCalData() As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean, myXlonCalOptn As Long
  Private myXlonCalDataRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonXCalRowCol As Long, myXlonYCalRowCol As Long
  Private myXlonCalStep As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double
  Private myXlonOrgXCnt As Long, myZdouOrgX() As Double
  Private myXlonOrgYCnt As Long, myZdouOrgY() As Double
  Private myXlonStrtNo As Long, myXlonLastNo As Long
  Private myXlonCalStrtNo As Long, myXlonCalLastNo As Long
  Private myXlonCalTmpCnt As Long, myZdouCalTmp() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgRowCnt = Empty: myXlonOrgColCnt = Empty: Erase myZdouOrgData
    myXlonOrgXCnt = Empty: Erase myZdouOrgX
    myXlonOrgYCnt = Empty: Erase myZdouOrgY
    myXlonStrtNo = Empty: myXlonLastNo = Empty
    myXlonCalStrtNo = Empty: myXlonCalLastNo = Empty
    myXlonCalTmpCnt = Empty: Erase myZdouCalTmp
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getStrtRowCol() As Long
    getStrtRowCol = myXlonStrtRowCol
End Property

Public Property Get getLastRowCol() As Long
    getLastRowCol = myXlonLastRowCol
End Property

Public Property Get getCalRowCnt() As Long
    getCalRowCnt = myXlonCalRowCnt
End Property

Public Property Get getCalColCnt() As Long
    getCalColCnt = myXlonCalColCnt
End Property

Public Property Get getCalData(ByVal i As Long, ByVal j As Long) As Double
    getCalData = myZdouCalData(i, j)
'myZdouCalData(i, 1) or myZdouCalData(1, j) : 算出データ配列
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonStrtRowCol = Empty: myXlonLastRowCol = Empty
    myXlonCalRowCnt = Empty: myXlonCalColCnt = Empty: Erase myZdouCalData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letRowDrctn(ByVal RHS As Boolean)
    myXbisRowDrctn = RHS
'myXbisRowDrctn = True  : 行方向のデータを処理
'myXbisRowDrctn = False : 列方向のデータを処理
End Property

Public Property Let letCalOptn(ByVal RHS As Long)
    myXlonCalOptn = RHS
'myXlonCalOptn = 1 : 回帰直線のY軸データ
'myXlonCalOptn = 2 : 回帰直線のX軸データ
'myXlonCalOptn = 3 : Y軸データと回帰直線との差分データ
'myXlonCalOptn = 4 : X軸データと回帰直線との差分データ
End Property

Public Property Let letCalDataRngOptn(ByVal RHS As Long)
    myXlonCalDataRngOptn = RHS
'myXlonCalDataRngOptn = 1 : 全データ範囲で取得
'myXlonCalDataRngOptn = 2 : 指定データ範囲で取得
End Property

Public Property Let letOrgData(ByRef RHS() As Double)
'myZdouOrgDataIN(i, j) : 元データ配列
    myZdouOrgDataIN = RHS
End Property

Public Property Let letXCalRowCol(ByVal RHS As Long)
    myXlonXCalRowCol = RHS
End Property

Public Property Let letYCalRowCol(ByVal RHS As Long)
    myXlonYCalRowCol = RHS
End Property

Public Property Let letCalStep(ByVal RHS As Long)
    myXlonCalStep = RHS
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
    
    If myXlonCalOptn < 1 Or myXlonCalOptn > 4 Then GoTo ExitPath
    If myXlonCalDataRngOptn < 1 Or myXlonCalDataRngOptn > 2 Then myXlonCalDataRngOptn = 1
    
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
    
    If myXbisRowDrctn = True Then
        If myXlonBgnRowCol < LBound(myZdouOrgData, 1) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 1)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 1) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 1)
        If myXlonXCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonXCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonYCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonYCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonCalStep <= 0 Or myXlonCalStep > UBound(myZdouOrgData, 1) Then _
            myXlonCalStep = 1
    Else
        If myXlonBgnRowCol < LBound(myZdouOrgData, 2) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 2)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 2) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 2)
        If myXlonXCalRowCol < LBound(myZdouOrgData, 1) _
            Or myXlonXCalRowCol > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonYCalRowCol < LBound(myZdouOrgData, 1) _
            Or myXlonYCalRowCol > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonCalStep <= 0 Or myXlonCalStep > UBound(myZdouOrgData, 2) Then _
            myXlonCalStep = 1
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
    
'//C:データ用変数を設定
    Call setDataVariables
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:計算実行
    Select Case myXlonCalOptn
        Case 1: Call prsCalRegressionLineYArray
        Case 2: Call prsCalRegressionLineXArray
        Case 3: Call prsCalDeviationYArrayFromRegressionLine
        Case 4: Call prsCalDeviationXArrayFromRegressionLine
        Case Else
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//C:出力データを格納
    Call setOutputVariables
            
ExitPath:
    Call initializeModuleVariables
End Sub

'SetP_データ用変数を設定する
Private Sub setDataVariables()
    myXbisExitFlag = False
    
  Dim myXlonBgnNo As Long, myXlonEndNo As Long
    If myXbisRowDrctn = True Then
        myXlonBgnNo = LBound(myZdouOrgData, 1)
        myXlonEndNo = UBound(myZdouOrgData, 1)
        Call PfixArrayPartColumnNumeric(myXlonOrgXCnt, myZdouOrgX, _
                myZdouOrgData, myXlonXCalRowCol, _
                myXlonBgnNo, myXlonEndNo, myXlonCalStep)
        Call PfixArrayPartColumnNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonYCalRowCol, _
                myXlonBgnNo, myXlonEndNo, myXlonCalStep)
    Else
        myXlonBgnNo = LBound(myZdouOrgData, 2)
        myXlonEndNo = UBound(myZdouOrgData, 2)
        Call PfixArrayPartRowNumeric(myXlonOrgXCnt, myZdouOrgX, _
                myZdouOrgData, myXlonXCalRowCol, _
                myXlonBgnNo, myXlonEndNo, myXlonCalStep)
        Call PfixArrayPartRowNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonYCalRowCol, _
                myXlonBgnNo, myXlonEndNo, myXlonCalStep)
    End If
    If myXlonOrgXCnt <= 0 Or myXlonOrgYCnt <= 0 Then GoTo ExitPath
    If myXlonOrgXCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
    myXlonStrtNo = myXlonBgnNo
    myXlonLastNo = myXlonEndNo
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:回帰直線のY軸データ
Private Sub prsCalRegressionLineYArray()
    myXbisExitFlag = False
    
'//データ配列の回帰直線を導出
  Dim myXlonRgrsnLnCnt As Long, myZdouRgrsnLnData() As Double
    Call incPfixRegressionLineYArray( _
            myXlonRgrsnLnCnt, myZdouRgrsnLnData, _
            myZdouOrgX, myZdouOrgY)
    If myXlonRgrsnLnCnt <= 0 Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
  Dim myXlonCalStrtNo As Long, myXlonCalLastNo As Long
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
        Case Else
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
    End Select
    
'//データを配列に格納
  Dim i As Long, n As Long: n = 0
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double
    
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myZdouCalTmp(n) = myZdouRgrsnLnData(i)
    Next i
    
    Erase myZdouRgrsnLnData
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:回帰直線のX軸データ
Private Sub prsCalRegressionLineXArray()
    myXbisExitFlag = False
    
'//データ配列の回帰直線を導出
  Dim myXlonRgrsnLnCnt As Long, myZdouRgrsnLnData() As Double
    Call incPfixRegressionLineXArray( _
            myXlonRgrsnLnCnt, myZdouRgrsnLnData, _
            myZdouOrgX, myZdouOrgY)
    If myXlonRgrsnLnCnt <= 0 Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
        Case Else
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
    End Select
    
'//データを配列に格納
  Dim i As Long, n As Long: n = 0
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double
    
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myZdouCalTmp(n) = myZdouRgrsnLnData(i)
    Next i
    
    Erase myZdouRgrsnLnData
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:Y軸データと回帰直線との差分データ
Private Sub prsCalDeviationYArrayFromRegressionLine()
    myXbisExitFlag = False
    
'//データ配列の回帰直線を導出
  Dim myXlonRgrsnLnCnt As Long, myZdouRgrsnLnData() As Double
    Call incPfixRegressionLineYArray( _
            myXlonRgrsnLnCnt, myZdouRgrsnLnData, _
            myZdouOrgX, myZdouOrgY)
    If myXlonRgrsnLnCnt <= 0 Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
        Case Else
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
    End Select
    
'//Y軸データと回帰直線との差分を導出
  Dim i As Long, n As Long: n = 0
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double
    
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myZdouCalTmp(n) = myZdouOrgY(i) - myZdouRgrsnLnData(i)
    Next i
    
    Erase myZdouRgrsnLnData
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:X軸データと回帰直線との差分データ
Private Sub prsCalDeviationXArrayFromRegressionLine()
    myXbisExitFlag = False
    
'//データ配列の回帰直線を導出
  Dim myXlonRgrsnLnCnt As Long, myZdouRgrsnLnData() As Double
    Call incPfixRegressionLineXArray( _
            myXlonRgrsnLnCnt, myZdouRgrsnLnData, _
            myZdouOrgX, myZdouOrgY)
    If myXlonRgrsnLnCnt <= 0 Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
        Case Else
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
    End Select
    
'//X軸データと回帰直線との差分を導出
  Dim i As Long, n As Long: n = 0
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double
    
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myZdouCalTmp(n) = myZdouOrgX(i) - myZdouRgrsnLnData(i)
    Next i
    
    Erase myZdouRgrsnLnData
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_出力データを格納する
Private Sub setOutputVariables()
    myXbisExitFlag = False
    
    myXlonStrtRowCol = myXlonCalStrtNo
    myXlonLastRowCol = myXlonCalLastNo
    
  Dim i As Long, j As Long
    If myXbisRowDrctn = True Then
        myXlonCalRowCnt = myXlonCalTmpCnt
        myXlonCalColCnt = 1
        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
        For i = LBound(myZdouCalTmp) To UBound(myZdouCalTmp)
            myZdouCalData(i, 1) = myZdouCalTmp(i)
        Next i
    Else
        myXlonCalRowCnt = 1
        myXlonCalColCnt = myXlonCalTmpCnt
        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
        For j = LBound(myZdouCalTmp) To UBound(myZdouCalTmp)
            myZdouCalData(1, j) = myZdouCalTmp(j)
        Next j
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
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

 '定型Ｐ_2次元の数値配列の指定行から一部を1次元配列で取得する
Private Sub PfixArrayPartRowNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, ByVal myXlonExtrctRow As Long, _
            ByVal myXlonBgnCol As Long, ByVal myXlonEndCol As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i)   : 抽出データ配列
'myZdouOrgData(i, j) : 元データ配列
    myXlonExtrctCnt = Empty: Erase myZdouExtrctData
    If myXlonBgnCol <= 0 Then Exit Sub
    If myXlonEndCol < myXlonBgnCol Then Exit Sub
    On Error GoTo ExitPath
    If myXlonEndCol > UBound(myZdouOrgData, 2) Then Exit Sub
    If myXlonExtrctRow < LBound(myZdouOrgData, 1) _
        Or myXlonExtrctRow > UBound(myZdouOrgData, 1) Then Exit Sub
    If coXlonStep <= 0 Or coXlonStep > UBound(myZdouOrgData, 2) Then Exit Sub
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(myXlonExtrctRow, myXlonBgnCol)
    On Error GoTo 0
    myXlonExtrctCnt = Int((myXlonEndCol - myXlonBgnCol + 0.9) / coXlonStep) + 1
    ReDim myZdouExtrctData(myXlonExtrctCnt) As Double
  Dim j As Long, n As Long: n = 0
    For j = myXlonBgnCol To myXlonEndCol Step coXlonStep
        n = n + 1: myZdouExtrctData(n) = myZdouOrgData(myXlonExtrctRow, j)
    Next j
ExitPath:
End Sub

 '定型Ｐ_2次元の数値配列の指定列から一部を1次元配列で取得する
Private Sub PfixArrayPartColumnNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, ByVal myXlonExtrctCol As Long, _
            ByVal myXlonBgnRow As Long, ByVal myXlonEndRow As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i)   : 抽出データ配列
'myZdouOrgData(i, j) : 元データ配列
    myXlonExtrctCnt = Empty: Erase myZdouExtrctData
    If myXlonBgnRow <= 0 Then Exit Sub
    If myXlonEndRow < myXlonBgnRow Then Exit Sub
    On Error GoTo ExitPath
    If myXlonEndRow > UBound(myZdouOrgData, 1) Then Exit Sub
    If myXlonExtrctCol < LBound(myZdouOrgData, 2) _
        Or myXlonExtrctCol > UBound(myZdouOrgData, 2) Then Exit Sub
    If coXlonStep <= 0 Or coXlonStep > UBound(myZdouOrgData, 1) Then Exit Sub
  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(myXlonBgnRow, myXlonExtrctCol)
    On Error GoTo 0
    myXlonExtrctCnt = Int((myXlonEndRow - myXlonBgnRow + 0.9) / coXlonStep) + 1
    ReDim myZdouExtrctData(myXlonExtrctCnt) As Double
  Dim i As Long, n As Long: n = 0
    For i = myXlonBgnRow To myXlonEndRow Step coXlonStep
        n = n + 1: myZdouExtrctData(n) = myZdouOrgData(i, myXlonExtrctCol)
    Next i
ExitPath:
End Sub

 '定型Ｐ_配列計算-回帰直線のY軸データ配列
Private Sub incPfixRegressionLineYArray( _
            myXlonRgrsnLnYCnt As Long, myZdouRgrsnLnYData() As Double, _
            ByRef myZdouOrgX() As Double, ByRef myZdouOrgY() As Double)
'Includes PfixSlopeIntercept
'myZdouRgrsnLnYData(i) : 回帰直線のY軸データ配列
'myZdouOrgX(i) : X軸の元データ配列
'myZdouOrgY(i) : Y軸の元データ配列
    myXlonRgrsnLnYCnt = Empty: Erase myZdouRgrsnLnYData
  Dim myXdouSlp As Double, myXdouIntcpt As Double
    Call PfixSlopeIntercept(myXdouSlp, myXdouIntcpt, myZdouOrgX, myZdouOrgY)
  Dim i As Long, U As Long, L As Long
    U = UBound(myZdouOrgX): L = LBound(myZdouOrgX)
    ReDim myZdouRgrsnLnYData(U) As Double
    myXlonRgrsnLnYCnt = U - L + 1
    For i = L To U
        myZdouRgrsnLnYData(i) = myXdouSlp * myZdouOrgX(i) + myXdouIntcpt
    Next i
End Sub

 '定型Ｐ_配列計算-回帰直線のX軸データ配列
Private Sub incPfixRegressionLineXArray( _
            myXlonRgrsnLnXCnt As Long, myZdouRgrsnLnXData() As Double, _
            ByRef myZdouOrgX() As Double, ByRef myZdouOrgY() As Double)
'Includes PfixSlopeIntercept
'myZdouRgrsnLnXData(i) : 回帰直線のX軸データ配列
'myZdouOrgX(i) : X軸の元データ配列
'myZdouOrgY(i) : Y軸の元データ配列
    myXlonRgrsnLnXCnt = Empty: Erase myZdouRgrsnLnXData
  Dim myXdouSlp As Double, myXdouIntcpt As Double
    Call PfixSlopeIntercept(myXdouSlp, myXdouIntcpt, myZdouOrgX, myZdouOrgY)
    If myXdouSlp = 0 Then Exit Sub
  Dim i As Long, U As Long, L As Long
    U = LBound(myZdouOrgX): L = UBound(myZdouOrgX)
    ReDim myZdouRgrsnLnXData(U) As Double
    myXlonRgrsnLnXCnt = U - L + 1
    For i = L To U
        myZdouRgrsnLnXData(i) = (myZdouOrgY(i) - myXdouIntcpt) / myXdouSlp
    Next i
End Sub

 '定型Ｐ_配列計算-2配列の傾き＆切片
Private Sub PfixSlopeIntercept( _
            myXdouSlp As Double, myXdouIntcpt As Double, _
            ByRef myZdouOrgX() As Double, ByRef myZdouOrgY() As Double)
'myZdouOrgX(i) : X軸の元データ配列
'myZdouOrgY(i) : Y軸の元データ配列
    myXdouSlp = Empty: myXdouIntcpt = Empty
    If UBound(myZdouOrgX) <> UBound(myZdouOrgY) Then Exit Sub
  Dim n As Long: n = UBound(myZdouOrgX) - LBound(myZdouOrgX) + 1
  Dim i As Long, SY2 As Double, SX2 As Double, SY As Double, SXY As Double, SX As Double
    For i = LBound(myZdouOrgX) To UBound(myZdouOrgX)
        SY2 = SY2 + (myZdouOrgY(i)) ^ 2
        SX2 = SX2 + (myZdouOrgX(i)) ^ 2
        SY = SY + myZdouOrgY(i)
        SXY = SXY + myZdouOrgX(i) * myZdouOrgY(i)
        SX = SX + myZdouOrgX(i)
    Next i
    myXdouSlp = (n * SXY - SY * SX) / (n * SX2 - SX ^ 2)
    myXdouIntcpt = (SX2 * SY - SXY * SX) / (n * SX2 - SX ^ 2)
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_データ配列の全区間で回帰直線処理をしてデータを取得する
'Private Sub instCRgrsnLn()
'  Const coXlonCalOptn As Long = 1
'    'coXlonCalOptn = 1 : 回帰直線のY軸データ
'    'coXlonCalOptn = 2 : 回帰直線のX軸データ
'    'coXlonCalOptn = 3 : Y軸データと回帰直線との差分データ
'    'coXlonCalOptn = 4 : X軸データと回帰直線との差分データ
'  Const coXlonCalDataRngOptn As Long = 1
'    'coXlonCalDataRngOptn = 1 : 全データ範囲で取得
'    'coXlonCalDataRngOptn = 2 : 指定データ範囲で取得
'  Const coXlonXCalRowCol As Long = 1
'  Const coXlonYCalRowCol As Long = 5
'  Const coXlonCalStep As Long = 1
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
''  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
''    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
''    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
''  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
''  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
''//CRgrsnLnクラスモジュールのインスタンス生成
'  Dim myXinsRgrsnLn As CRgrsnLn: Set myXinsRgrsnLn = New CRgrsnLn
'  Dim i As Long, j As Long
'    With myXinsRgrsnLn
'    '//クラス内変数への入力
'        .letRowDrctn = True
'        .letCalOptn = coXlonCalOptn
'        .letCalDataRngOptn = coXlonCalDataRngOptn
'        .letOrgData = myZdouOrgData
'        .letXCalRowCol = coXlonXCalRowCol
'        .letYCalRowCol = coXlonYCalRowCol
'        .letCalStep = coXlonCalStep
'        .letBgnRowCol = coXlonBgnRowCol
'        .letEndRowCol = coXlonEndRowCol
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonStrtRowCol = .getStrtRowCol
'        myXlonLastRowCol = .getLastRowCol
'        myXlonCalRowCnt = .getCalRowCnt
'        myXlonCalColCnt = .getCalColCnt
'        If myXlonCalRowCnt <= 0 Or myXlonCalColCnt <= 0 Then GoTo JumpPath
'        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
'        For j = 1 To myXlonCalColCnt
'            For i = 1 To myXlonCalRowCnt
'                myZdouCalData(i, j) = .getCalData(i, j)
'            Next i
'        Next j
'    End With
'JumpPath:
'    Set myXinsRgrsnLn = Nothing
'    Call variablesOfCRgrsnLn(myXlonCalRowCnt, myZdouCalData)   'Debug.Print
'End Sub
'Private Sub variablesOfCRgrsnLn( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CRgrsnLnクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
