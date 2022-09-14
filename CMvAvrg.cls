VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMvAvrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes PfixArrayPartColumnNumeric
'Includes PfixArrayPartRowNumeric
'Includes Pfix1DArrayExtractNumeric
'Includes PfixTrapezoidWaveLong
'Includes PfncdouArithmeticMean
'Includes PfncdouGeometricMean
'Includes PfncdouMovingAverageFt

Option Explicit
Option Base 1

'◆ClassProc名_データ配列を移動平均処理する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CMvAvrg"
  
'//PropertyGet_出力データ
  Private myXlonStrtRowCol As Long, myXlonLastRowCol As Long
  Private myXlonCalRowCnt As Long, myXlonCalColCnt As Long, _
            myZdouCalData() As Double
  Private myXdouFt As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean, myXlonCalOptn As Long, myXlonCalRng As Long
  Private myXlonCalDataRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonCalRowCol As Long
  Private myXlonCalSmplQnt As Long, myXlonCalStep As Long
  Private myXdouTs As Double
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double
  Private myXlonOrgTmpCnt As Long, myZdouOrgTmp() As Double
  Private myXlonStrtNo As Long, myXlonLastNo As Long
  Private myXlonPreOfst As Long, myXlonPstOfst As Long
  Private myXlonCalStrtNo As Long, myXlonCalLastNo As Long
  Private myXlonCalTmpCnt As Long, myZdouCalTmp() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgRowCnt = Empty: myXlonOrgColCnt = Empty: Erase myZdouOrgData
    myXlonOrgTmpCnt = Empty: Erase myZdouOrgTmp
    myXlonStrtNo = Empty: myXlonLastNo = Empty
    myXlonPreOfst = Empty: myXlonPstOfst = Empty
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
'myZdouCalData(i, j) : 移動平均後データ配列
End Property

Public Property Get getFt() As Double
    getFt = myXdouFt
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonStrtRowCol = Empty: myXlonLastRowCol = Empty
    myXlonCalRowCnt = Empty: myXlonCalColCnt = Empty: Erase myZdouCalData
    myXdouFt = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letRowDrctn(ByVal RHS As Boolean)
    myXbisRowDrctn = RHS
'myXbisRowDrctn = True  : 行方向のデータを処理
'myXbisRowDrctn = False : 列方向のデータを処理
End Property

Public Property Let letCalOptn(ByVal RHS As Long)
    myXlonCalOptn = RHS
'myXlonCalOptn = 1 : 相加平均値
'myXlonCalOptn = 2 : 相乗平均値
End Property

Public Property Let letCalRng(ByVal RHS As Long)
    myXlonCalRng = RHS
'myXlonCalRng = -1 : 今回値より前のデータで計算
'myXlonCalRng = 0  : 今回値の前後のデータで計算
'myXlonCalRng = 1  : 今回値より後のデータで計算
End Property

Public Property Let letCalDataRngOptn(ByVal RHS As Long)
    myXlonCalDataRngOptn = RHS
'myXlonCalDataRngOptn = 1 : 全データ範囲で取得
'myXlonCalDataRngOptn = 2 : 指定データ範囲で取得
'myXlonCalDataRngOptn = 3 : 有効データ範囲で取得
End Property

Public Property Let letOrgData(ByRef RHS() As Double)
'myZdouOrgDataIN(i, j) : 元データ配列
    myZdouOrgDataIN = RHS
End Property

Public Property Let letCalRowCol(ByVal RHS As Long)
    myXlonCalRowCol = RHS
End Property

Public Property Let letCalSmplQnt(ByVal RHS As Long)
    myXlonCalSmplQnt = RHS
End Property

Public Property Let letCalStep(ByVal RHS As Long)
    myXlonCalStep = RHS
End Property

Public Property Let letTs(ByVal RHS As Double)
    myXdouTs = RHS
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
    
    If myXlonCalOptn < 1 Or myXlonCalOptn > 2 Then myXlonCalOptn = 1
    If myXlonCalRng < -1 Or myXlonCalRng > 1 Then myXlonCalRng = 0
    If myXlonCalDataRngOptn < 1 Or myXlonCalDataRngOptn > 3 Then myXlonCalDataRngOptn = 1
    
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
        If myXlonCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonCalSmplQnt <= 0 _
            Or myXlonCalSmplQnt > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonCalStep <= 0 Or myXlonCalStep > UBound(myZdouOrgData, 1) Then _
            myXlonCalStep = 1
        If myXlonBgnRowCol < LBound(myZdouOrgData, 1) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 1)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 1) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 1)
    Else
        If myXlonCalRowCol < LBound(myZdouOrgData, 1) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonCalSmplQnt <= 0 _
            Or myXlonCalSmplQnt > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonCalStep <= 0 Or myXlonCalStep > UBound(myZdouOrgData, 2) Then _
            myXlonCalStep = 1
        If myXlonBgnRowCol < LBound(myZdouOrgData, 2) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 2)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 2) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 2)
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
    
'//C:制御用変数を設定
    Call setControlVariables
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag

'//P:計算実行:移動平均
    Select Case myXlonCalOptn
        Case 1: Call prsCalArithmeticMean
        Case 2: Call prsCalGeometricMean
        Case Else
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//C:出力データを格納
    Call setOutputVariables
        
'//単純移動平均処理のカットオフ周波数を導出
    If myXlonCalOptn = 1 Then _
        myXdouFt = PfncdouMovingAverageFt(myXdouTs * myXlonCalStep, myXlonCalSmplQnt)
    
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
        Call PfixArrayPartColumnNumeric(myXlonOrgTmpCnt, myZdouOrgTmp, _
                myZdouOrgData, myXlonCalRowCol, myXlonBgnNo, myXlonEndNo)
    Else
        myXlonBgnNo = LBound(myZdouOrgData, 2)
        myXlonEndNo = UBound(myZdouOrgData, 2)
        Call PfixArrayPartRowNumeric(myXlonOrgTmpCnt, myZdouOrgTmp, _
                myZdouOrgData, myXlonCalRowCol, myXlonBgnNo, myXlonEndNo)
    End If
    If myXlonOrgTmpCnt <= 0 Then GoTo ExitPath
    
    myXlonStrtNo = myXlonBgnNo
    myXlonLastNo = myXlonEndNo
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
    myXbisExitFlag = False
    
'//計算するオフセット量を設定
    Select Case myXlonCalRng
        Case -1
            myXlonPreOfst = (myXlonCalSmplQnt - 1) * myXlonCalStep
            myXlonPstOfst = 0
        Case 1
            myXlonPreOfst = 0
            myXlonPstOfst = (myXlonCalSmplQnt - 1) * myXlonCalStep
        Case Else
            myXlonPreOfst = Int(myXlonCalSmplQnt / 2) * myXlonCalStep
            myXlonPstOfst = Int(myXlonCalSmplQnt / 2) * myXlonCalStep
    End Select
    
'//計算が有効な最初と最後の位置を設定
    myXlonStrtNo = myXlonStrtNo + myXlonPreOfst
    myXlonLastNo = myXlonLastNo - myXlonPstOfst
    If myXlonLastNo < myXlonStrtNo Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:移動平均_相加平均
Private Sub prsCalArithmeticMean()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = LBound(myZdouOrgTmp)
            myXlonCalLastNo = UBound(myZdouOrgTmp)
        Case 2
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
        Case Else
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
    End Select
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgTmpCnt, myXlonStrtNo, myXlonLastNo, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgTmpCnt Then GoTo ExitPath
    
'//計算処理:移動平均_相加平均
  Dim i As Long, n As Long: n = 0
  Dim myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double

    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
        myXlonEndNo = i + myXlonPstOfst + myZlonAdjstVal(i)
        Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                myZdouOrgTmp, myXlonBgnNo, myXlonEndNo, myXlonCalStep)
        If myXlonTmpCnt <= 0 Then GoTo ExitPath
        myZdouCalTmp(n) = PfncdouArithmeticMean(myZdouTmp)
    Next i
    If n <> myXlonCalTmpCnt Then GoTo ExitPath
    
    Erase myZlonAdjstVal: Erase myZdouTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:移動平均_相乗平均
Private Sub prsCalGeometricMean()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = LBound(myZdouOrgTmp)
            myXlonCalLastNo = UBound(myZdouOrgTmp)
        Case 2
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
        Case Else
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
    End Select
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgTmpCnt, myXlonStrtNo, myXlonLastNo, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgTmpCnt Then GoTo ExitPath
    
'//計算処理:移動平均_相乗平均
  Dim i As Long, n As Long: n = 0
  Dim myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double

    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
        myXlonEndNo = i + myXlonPstOfst + myZlonAdjstVal(i)
        Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                myZdouOrgTmp, myXlonBgnNo, myXlonEndNo, myXlonCalStep)
        If myXlonTmpCnt <= 0 Then GoTo ExitPath
        myZdouCalTmp(i) = PfncdouGeometricMean(myZdouTmp)
    Next i
    If n <> myXlonCalTmpCnt Then GoTo ExitPath
    
    Erase myZlonAdjstVal: Erase myZdouTmp
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

 '定型Ｐ_2次元の数値配列の指定列から一部を1次元配列で取得する
Private Sub PfixArrayPartColumnNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, ByVal myXlonExtrctCol As Long, _
            ByVal myXlonBgnRow As Long, ByVal myXlonEndRow As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i)   : 抽出データ配列
'myZdouOrgData(i, j) : 元データ配列
    myXlonExtrctCnt = Empty: Erase myZdouExtrctData
    If myXlonBgnRow < LBound(myZdouOrgData, 2) Then Exit Sub
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

 '定型Ｐ_2次元の数値配列の指定行から一部を1次元配列で取得する
Private Sub PfixArrayPartRowNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, ByVal myXlonExtrctRow As Long, _
            ByVal myXlonBgnCol As Long, ByVal myXlonEndCol As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i)   : 抽出データ配列
'myZdouOrgData(i, j) : 元データ配列
    myXlonExtrctCnt = Empty: Erase myZdouExtrctData
    If myXlonBgnCol < LBound(myZdouOrgData, 1) Then Exit Sub
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

 '定型Ｐ_1次元の数値配列から一部を取得する
Private Sub Pfix1DArrayExtractNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, _
            ByVal myXlonBgn As Long, ByVal myXlonEnd As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i): 取得データ配列
'myZdouOrgData(i): 元データ配列
    myXlonExtrctCnt = Empty: Erase myZdouExtrctData
  Dim myXdouTmp As Double, L As Long
    On Error GoTo ExitPath
    If myXlonBgn < LBound(myZdouOrgData) Then Exit Sub
    If myXlonEnd > UBound(myZdouOrgData) Then Exit Sub
    If myXlonEnd < myXlonBgn Then Exit Sub
    If coXlonStep <= 0 Or coXlonStep > UBound(myZdouOrgData) Then Exit Sub
    myXdouTmp = myZdouOrgData(myXlonBgn): L = LBound(myZdouOrgData)
    On Error GoTo 0
  Dim i As Long, n As Long: n = L - 1
    myXlonExtrctCnt = Int((myXlonEnd - myXlonBgn + 0.9) / coXlonStep) + 1
    i = myXlonExtrctCnt - (1 - L): ReDim myZdouExtrctData(i) As Double
    For i = myXlonBgn To myXlonEnd Step coXlonStep
        n = n + 1: myZdouExtrctData(n) = myZdouOrgData(i)
    Next i
ExitPath:
End Sub

 '定型Ｐ_整数で台形波形を生成する
Private Sub PfixTrapezoidWaveLong( _
            myXlonCalDataCnt As Long, myZlonCalData() As Long, _
            ByVal myXlonDataQnt As Long, _
            ByVal myXlonInflctPntBgn As Long, ByVal myXlonInflctPntEnd As Long, _
            Optional ByVal coXlonStrtStp As Long = -1, _
            Optional ByVal coXlonLastStp As Long = 1, _
            Optional ByVal coXlonCnstVal As Long = 0)
'myZlonCalData(i) : 取得データ配列
'myXlonInflctPntBgn : 開始側の変曲点
'myXlonInflctPntEnd : 終了側の変曲点
    myXlonCalDataCnt = Empty:  Erase myZlonCalData
    If myXlonDataQnt <= 0 Then Exit Sub
    If myXlonInflctPntBgn < 1 Then Exit Sub
    If myXlonInflctPntEnd > myXlonDataQnt Then Exit Sub
    If myXlonInflctPntBgn > myXlonInflctPntEnd Then Exit Sub
    On Error GoTo ExitPath
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    Erase myZlonTmp
    On Error GoTo 0
  Dim i As Long, n As Long
    myXlonCalDataCnt = myXlonDataQnt
    i = myXlonCalDataCnt - (1 - L): ReDim myZlonCalData(i) As Long
    If myXlonInflctPntBgn = 1 Then GoTo JumpPath1
    n = myXlonInflctPntBgn
    For i = 1 To myXlonInflctPntBgn - 1
        n = n - 1
        myZlonCalData(i) = coXlonCnstVal - coXlonStrtStp * n
    Next i
JumpPath1:
    If coXlonCnstVal = 0 Then GoTo JumpPath2
    For i = myXlonInflctPntBgn To myXlonInflctPntEnd
        myZlonCalData(i) = coXlonCnstVal
    Next i
JumpPath2:
    If myXlonInflctPntEnd = 1 Then Exit Sub
    n = 0
    For i = myXlonInflctPntEnd + 1 To myXlonDataQnt
        n = n + 1
        myZlonCalData(i) = coXlonCnstVal + coXlonLastStp * n
    Next i
ExitPath:
End Sub

 '定型Ｆ_配列計算-相加平均
Private Function PfncdouArithmeticMean(ByRef myZdouOrgData() As Double) As Double
'myZdouOrgData(i) : 元データ配列
    PfncdouArithmeticMean = Empty
  Dim myXdouAveTmp As Double: myXdouAveTmp = 0
  Dim i As Long, n As Long: n = 0
    For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
        n = n + 1: myXdouAveTmp = myXdouAveTmp + myZdouOrgData(i)
    Next i
    PfncdouArithmeticMean = myXdouAveTmp / n
End Function

 '定型Ｆ_配列計算-相乗平均
Private Function PfncdouGeometricMean(ByRef myZdouOrgData() As Double) As Double
'myZdouOrgData(i) : 元データ配列
    PfncdouGeometricMean = Empty
  Dim myXdouAveTmp As Double: myXdouAveTmp = 1
  Dim i As Long, n As Long: n = 0
    For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
        n = n + 1: myXdouAveTmp = myXdouAveTmp * myZdouOrgData(i)
    Next i
    PfncdouGeometricMean = myXdouAveTmp ^ (1 / n)
End Function

 '定型Ｆ_単純移動平均処理のカットオフ周波数を導出する
Private Function PfncdouMovingAverageFt( _
            ByVal myXdouTs As Double, ByVal myXlonM As Long) As Double
'myXdouTs : サンプルリング時間[sec]
'myXdouM  : 移動平均点数
    PfncdouMovingAverageFt = Empty
    If myXdouTs <= 0 Then Exit Function
    If myXlonM <= 0 Then Exit Function
    PfncdouMovingAverageFt = 0.443 / myXdouTs / myXlonM
End Function

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================
'■参考モジュール

' '定型Ｐ_データ配列の指定列を相加平均したデータ配列を取得する
'Private Sub incPfixVDataListArithmeticMean( _
'            myXlonStrtRow As Long, myXlonLastRow As Long, _
'            myXlonMvAveRowCnt As Long, myXlonMvAveColCnt As Long, _
'            myZdouMvAveData() As Double, _
'            ByRef myZdouOrgData() As Double, ByVal myXlonCalCol As Long, _
'            ByVal myXlonCalSmplQnt As Long, _
'            Optional ByVal coXlonBgnRow As Long = 1, _
'            Optional ByVal coXlonEndRow As Long = 0, _
'            Optional ByVal coXlonCalRng As Long = 0, _
'            Optional ByVal coXlonCalStep As Long = 1)
''Includes PfixArrayPartColumnNumeric
''Includes PfncdouArithmeticMean
''myZdouMvAveData(i, j) : 移動平均後データ配列
''myZdouOrgData(i, j) : 元データ配列
''coXlonCalRng = -1 : 今回値より前のデータで平均化
''coXlonCalRng = 0  : 今回値の前後のデータで平均化
''coXlonCalRng = 1  : 今回値より後のデータで平均化
'    myXlonStrtRow = Empty: myXlonLastRow = Empty
'    myXlonMvAveRowCnt = Empty: myXlonMvAveColCnt = Empty: Erase myZdouMvAveData
'    On Error GoTo ExitPath
'    If myXlonCalSmplQnt <= 0 Or myXlonCalSmplQnt > UBound(myZdouOrgData, 1) Then Exit Sub
'    If coXlonBgnRow < LBound(myZdouOrgData, 1) Then coXlonBgnRow = LBound(myZdouOrgData, 1)
'    If coXlonEndRow < coXlonBgnRow Or coXlonEndRow > UBound(myZdouOrgData, 1) Then _
'        coXlonEndRow = UBound(myZdouOrgData, 1)
'    If coXlonCalStep <= 0 Or coXlonCalStep > UBound(myZdouOrgData, 1) Then _
'        coXlonCalStep = 1
'  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(coXlonBgnRow, myXlonCalCol)
'    On Error GoTo 0
''//計算する最初のデータ位置と最後のデータ位置を設定
'  Dim myXlonPreOfst As Long, myXlonPstOfst As Long
'    If coXlonCalRng = -1 Then
'        myXlonPreOfst = (myXlonCalSmplQnt - 1) * coXlonCalStep
'        myXlonPstOfst = 0
'    ElseIf coXlonCalRng = 1 Then
'        myXlonPreOfst = 0
'        myXlonPstOfst = (myXlonCalSmplQnt - 1) * coXlonCalStep
'    Else
'        myXlonPreOfst = Int(myXlonCalSmplQnt / 2) * coXlonCalStep
'        myXlonPstOfst = Int(myXlonCalSmplQnt / 2) * coXlonCalStep
'    End If
''//出力する配列サイズを設定
'    On Error GoTo ExitPath
'    myXlonStrtRow = LBound(myZdouOrgData, 1) + myXlonPreOfst
'    myXlonLastRow = UBound(myZdouOrgData, 1) - myXlonPstOfst
'    If myXlonStrtRow < coXlonBgnRow Then myXlonStrtRow = coXlonBgnRow
'    If myXlonLastRow > coXlonEndRow Then myXlonLastRow = coXlonEndRow
'    myXlonMvAveRowCnt = myXlonLastRow - myXlonStrtRow + 1
'    myXlonMvAveColCnt = UBound(myZdouOrgData, 2)
'    ReDim myZdouMvAveData(myXlonMvAveRowCnt, myXlonMvAveColCnt) As Double
'    On Error GoTo 0
''//計算実行
'  Dim i As Long, j As Long, n As Long: n = 0
'  Dim myXlonBgnNo As Long, myXlonEndNo As Long
'  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
'    For j = LBound(myZdouOrgData, 2) To UBound(myZdouOrgData, 2)
'        n = 0
'        If j = myXlonCalCol Then
'            For i = myXlonStrtRow To myXlonLastRow
'                n = n + 1
'                myXlonBgnNo = i - myXlonPreOfst
'                myXlonEndNo = i + myXlonPstOfst
'                Call PfixArrayPartColumnNumeric(myXlonTmpCnt, myZdouTmp, _
'                        myZdouOrgData, myXlonCalCol, _
'                        myXlonBgnNo, myXlonEndNo, coXlonCalStep)
'                If myXlonTmpCnt > 0 Then _
'                    myZdouMvAveData(n, j) = PfncdouArithmeticMean(myZdouTmp)
'            Next i
'        Else
'            For i = myXlonStrtRow To myXlonLastRow
'                n = n + 1
'                myZdouMvAveData(n, j) = myZdouOrgData(i, j)
'            Next i
'        End If
'    Next j
'    Exit Sub
'    Erase myZdouTmp
'ExitPath:
'End Sub
'
' '定型Ｐ_データ配列の指定行を相加平均したデータ配列を取得する
'Private Sub incPfixHDataListArithmeticMean( _
'            myXlonStrtCol As Long, myXlonLastCol As Long, _
'            myXlonMvAveRowCnt As Long, myXlonMvAveColCnt As Long, _
'            myZdouMvAveData() As Double, _
'            ByRef myZdouOrgData() As Double, ByVal myXlonCalRow As Long, _
'            ByVal myXlonCalSmplQnt As Long, _
'            Optional ByVal coXlonBgnCol As Long = 1, _
'            Optional ByVal coXlonEndCol As Long = 0, _
'            Optional ByVal coXlonCalRng As Long = 0, _
'            Optional ByVal coXlonCalStep As Long = 1)
''Includes PfixArrayPartRowNumeric
''Includes PfncdouArithmeticMean
''myZdouMvAveData(i, j) : 移動平均後データ配列
''myZdouOrgData(i, j) : 元データ配列
''coXlonCalRng = -1 : 今回値より前のデータで平均化
''coXlonCalRng = 0  : 今回値の前後のデータで平均化
''coXlonCalRng = 1  : 今回値より後のデータで平均化
'  Const coXbisRowDrctn As Boolean = False
'    myXlonStrtCol = Empty: myXlonLastCol = Empty
'    myXlonMvAveRowCnt = Empty: myXlonMvAveColCnt = Empty: Erase myZdouMvAveData
'    On Error GoTo ExitPath
'    If myXlonCalSmplQnt <= 0 Or myXlonCalSmplQnt > UBound(myZdouOrgData, 2) Then Exit Sub
'    If coXlonBgnCol < LBound(myZdouOrgData, 2) Then coXlonBgnCol = LBound(myZdouOrgData, 2)
'    If coXlonEndCol < coXlonBgnCol Or coXlonEndCol > UBound(myZdouOrgData, 2) Then _
'        coXlonEndCol = UBound(myZdouOrgData, 2)
'    If coXlonCalStep <= 0 Or coXlonCalStep > UBound(myZdouOrgData, 2) Then _
'        coXlonCalStep = 1
'  Dim myXdouTmp As Double: myXdouTmp = myZdouOrgData(myXlonCalRow, coXlonBgnCol)
'    On Error GoTo 0
''//計算する最初のデータ位置と最後のデータ位置を設定
'  Dim myXlonPreOfst As Long, myXlonPstOfst As Long
'    If coXlonCalRng = -1 Then
'        myXlonPreOfst = (myXlonCalSmplQnt - 1) * coXlonCalStep
'        myXlonPstOfst = 0
'    ElseIf coXlonCalRng = 1 Then
'        myXlonPreOfst = 0
'        myXlonPstOfst = (myXlonCalSmplQnt - 1) * coXlonCalStep
'    Else
'        myXlonPreOfst = Int(myXlonCalSmplQnt / 2) * coXlonCalStep
'        myXlonPstOfst = Int(myXlonCalSmplQnt / 2) * coXlonCalStep
'    End If
''//出力する配列サイズを設定
'    On Error GoTo ExitPath
'    myXlonStrtCol = LBound(myZdouOrgData, 2) + myXlonPreOfst
'    myXlonLastCol = UBound(myZdouOrgData, 2) - myXlonPstOfst
'    If myXlonStrtCol < coXlonBgnCol Then myXlonStrtCol = coXlonBgnCol
'    If myXlonLastCol > coXlonEndCol Then myXlonLastCol = coXlonEndCol
'    myXlonMvAveRowCnt = UBound(myZdouOrgData, 1)
'    myXlonMvAveColCnt = myXlonLastCol - myXlonStrtCol + 1
'    ReDim myZdouMvAveData(myXlonMvAveRowCnt, myXlonMvAveColCnt) As Double
'    On Error GoTo 0
''//計算実行
'  Dim i As Long, j As Long, n As Long: n = 0
'  Dim myXlonBgnNo As Long, myXlonEndNo As Long
'  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
'    For i = LBound(myZdouOrgData, 1) To UBound(myZdouOrgData, 1)
'        n = 0
'        If i = myXlonCalRow Then
'            For j = myXlonStrtCol To myXlonLastCol
'                n = n + 1
'                myXlonBgnNo = j - myXlonPreOfst
'                myXlonEndNo = j + myXlonPstOfst
'                Call PfixArrayPartRowNumeric(myXlonTmpCnt, myZdouTmp, _
'                        myZdouOrgData, myXlonCalRow, _
'                        myXlonBgnNo, myXlonEndNo, coXlonCalStep)
'                If myXlonTmpCnt > 0 Then _
'                    myZdouMvAveData(i, n) = PfncdouArithmeticMean(myZdouTmp)
'            Next j
'        Else
'            For j = myXlonStrtCol To myXlonLastCol
'                n = n + 1
'                myZdouMvAveData(i, n) = myZdouOrgData(i, j)
'            Next j
'        End If
'    Next i
'    Exit Sub
'ExitPath:
'End Sub

'===============================================================================================

''◆ClassProc名_データ配列を移動平均処理する
'Private Sub instCMvAvrg()
'  Const coXlonCalOptn As Long = 1
'    'coXlonCalOptn = 1 : 相加平均値
'    'coXlonCalOptn = 2 : 相乗平均値
'  Const coXlonCalRng As Long = 0
'    'coXlonCalRng = -1 : 今回値より前のデータで平均化
'    'coXlonCalRng = 0  : 今回値の前後のデータで平均化
'    'coXlonCalRng = 1  : 今回値より後のデータで平均化
'  Const coXlonCalDataRngOptn As Long = 1
'    'coXlonCalDataRngOptn = 1 : 全データ範囲で取得
'    'coXlonCalDataRngOptn = 2 : 指定データ範囲で取得
'    'coXlonCalDataRngOptn = 3 : 有効データ範囲で取得
'  Const coXlonCalRowCol As Long = 5
'  Const coXlonCalSmplQnt As Long = 5
'  Const coXlonCalStep As Long = 1
'  Const coXdouTs As Double = 0.001
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
'    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
'    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
'  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
'  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
'  Dim myXdouFt As Double
''//CMvAvrgクラスモジュールのインスタンス生成
'  Dim myXinsMvAvrg As CMvAvrg: Set myXinsMvAvrg = New CMvAvrg
'  Dim i As Long, j As Long
'    With myXinsMvAvrg
'    '//クラス内変数への入力
'        .letRowDrctn = True
'        .letCalOptn = coXlonCalOptn
'        .letCalRng = coXlonCalRng
'        .letCalDataRngOptn = coXlonCalDataRngOptn
'        .letOrgData = myZdouOrgData
'        .letCalRowCol = coXlonCalRowCol
'        .letCalSmplQnt = coXlonCalSmplQnt
'        .letCalStep = coXlonCalStep
'        .letTs = coXdouTs
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
'        myXdouFt = .getFt
'    End With
'JumpPath:
'    Set myXinsMvAvrg = Nothing
''    Call variablesOfCMvAvrg(myXlonCalRowCnt, myZdouCalData)  'Debug.Print
'End Sub
'Private Sub variablesOfCMvAvrg( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CMvAvrgクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
