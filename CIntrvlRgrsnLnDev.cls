VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIntrvlRgrsnLnDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes PfixArrayPartRowNumeric
'Includes PfixArrayPartColumnNumeric
'Includes Pfix1DArrayExtractNumeric
'Includes PfixTrapezoidWaveLong
'Includes PfixSlopeIntercept

Option Explicit
Option Base 1

'◆ClassProc名_データ配列の指定区間ごとで回帰直線間の差分データを取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CIntrvlRgrsnLnDev"
  
'//PropertyGet_出力データ
  Private myXlonStrtRowCol As Long, myXlonLastRowCol As Long
  Private myXlonCalRowCnt As Long, myXlonCalColCnt As Long, _
            myZdouCalData() As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean
  Private myXlonCalDataRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonXCalRowCol As Long, myXlonYCalRowCol As Long
  Private myXlonCalSmplQnt As Long, myXlonCalStep As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double
  Private myXlonOrgXCnt As Long, myZdouOrgX() As Double
  Private myXlonOrgYCnt As Long, myZdouOrgY() As Double
  Private myXlonStrtNo As Long, myXlonLastNo As Long
  Private myXlonPreOfstNeg As Long, myXlonPstOfstNeg As Long
  Private myXlonPreOfstPos As Long, myXlonPstOfstPos As Long
  Private myXlonCalStrtNo As Long, myXlonCalLastNo As Long
  Private myXlonCalTmpCnt As Long, myZdouCalTmp() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgRowCnt = Empty: myXlonOrgColCnt = Empty: Erase myZdouOrgData
    myXlonOrgXCnt = Empty: Erase myZdouOrgX
    myXlonOrgYCnt = Empty: Erase myZdouOrgY
    myXlonStrtNo = Empty: myXlonLastNo = Empty
    myXlonPreOfstNeg = Empty: myXlonPstOfstNeg = Empty
    myXlonPreOfstPos = Empty: myXlonPstOfstPos = Empty
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

Public Property Let letXCalRowCol(ByVal RHS As Long)
    myXlonXCalRowCol = RHS
End Property

Public Property Let letYCalRowCol(ByVal RHS As Long)
    myXlonYCalRowCol = RHS
End Property

Public Property Let letCalSmplQnt(ByVal RHS As Long)
    myXlonCalSmplQnt = RHS
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
        If myXlonBgnRowCol < LBound(myZdouOrgData, 1) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 1)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 1) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 1)
        If myXlonXCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonXCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonYCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonYCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonCalSmplQnt <= 0 _
            Or myXlonCalSmplQnt > UBound(myZdouOrgData, 1) Then GoTo ExitPath
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
        If myXlonCalSmplQnt <= 0 _
            Or myXlonCalSmplQnt > UBound(myZdouOrgData, 2) Then GoTo ExitPath
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
    
'//C:制御用変数を設定
    Call setControlVariables
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag

'//P:計算実行:回帰直線間の差分
    Call prsCal
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
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
                myZdouOrgData, myXlonXCalRowCol, myXlonBgnNo, myXlonEndNo)
        Call PfixArrayPartColumnNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonYCalRowCol, myXlonBgnNo, myXlonEndNo)
    Else
        myXlonBgnNo = LBound(myZdouOrgData, 2)
        myXlonEndNo = UBound(myZdouOrgData, 2)
        Call PfixArrayPartRowNumeric(myXlonOrgXCnt, myZdouOrgX, _
                myZdouOrgData, myXlonXCalRowCol, myXlonBgnNo, myXlonEndNo)
        Call PfixArrayPartRowNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonYCalRowCol, myXlonBgnNo, myXlonEndNo)
    End If
    If myXlonOrgXCnt <= 0 Or myXlonOrgYCnt <= 0 Then GoTo ExitPath
    If myXlonOrgXCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
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
    myXlonPreOfstNeg = (myXlonCalSmplQnt - 1) * myXlonCalStep
    myXlonPstOfstNeg = 0
    
    myXlonPreOfstPos = 0
    myXlonPstOfstPos = (myXlonCalSmplQnt - 1) * myXlonCalStep
    
'//計算が有効な最初と最後の位置を設定
    myXlonStrtNo = myXlonStrtNo + myXlonPreOfstNeg
    myXlonLastNo = myXlonLastNo - myXlonPstOfstPos
    If myXlonLastNo < myXlonStrtNo Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:回帰直線間の差分
Private Sub prsCal()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = LBound(myZdouOrgX)
            myXlonCalLastNo = UBound(myZdouOrgX)
        Case 2
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
        Case Else
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
    End Select
    
'//台形波形を生成
  Dim myXlonAdjstNegCnt As Long, myZlonAdjstNeg() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstNegCnt, myZlonAdjstNeg, _
            myXlonOrgXCnt, myXlonStrtNo, UBound(myZdouOrgX), -1, 0, 0)
    If myXlonAdjstNegCnt <> myXlonOrgXCnt Then GoTo ExitPath
    
  Dim myXlonAdjstPosCnt As Long, myZlonAdjstPos() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstPosCnt, myZlonAdjstPos, _
            myXlonOrgXCnt, LBound(myZdouOrgX), myXlonLastNo, 0, -1, 0)
    If myXlonAdjstPosCnt <> myXlonOrgXCnt Then GoTo ExitPath
    
'//計算処理:回帰直線間の差分
  Dim i As Long, n As Long: n = 0
  Dim myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXdouSlpNeg As Double, myXdouIntrcptNeg As Double
  Dim myXdouSlpPos As Double, myXdouIntrcptPos As Double
  Dim myXdouRgrsnNeg As Double, myXdouRgrsnPos As Double
    
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double

    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
    '//今回値より前のデータで計算
        myXlonBgnNo = i - myXlonPreOfstNeg + myZlonAdjstNeg(i)
        myXlonEndNo = i + myXlonPstOfstNeg
        Call CsubCalRegressionLineSlopeIntercept( _
                myXdouSlpNeg, myXdouIntrcptNeg, _
                myXlonBgnNo, myXlonEndNo)
        If myXdouSlpNeg = 0 And myXdouIntrcptNeg = 0 Then GoTo ExitPath
    '//今回値より後のデータで計算
        myXlonBgnNo = i - myXlonPreOfstPos
        myXlonEndNo = i + myXlonPstOfstPos + myZlonAdjstPos(i)
        Call CsubCalRegressionLineSlopeIntercept( _
                myXdouSlpPos, myXdouIntrcptPos, _
                myXlonBgnNo, myXlonEndNo)
        If myXdouSlpPos = 0 And myXdouIntrcptPos = 0 Then GoTo ExitPath
    '//回帰直線間の差分を計算
        myXdouRgrsnNeg = myXdouSlpNeg * myZdouOrgX(i) + myXdouIntrcptNeg
        myXdouRgrsnPos = myXdouSlpPos * myZdouOrgX(i) + myXdouIntrcptPos
        myZdouCalTmp(n) = myXdouRgrsnPos - myXdouRgrsnNeg
    Next i
    
    Erase myZlonAdjstNeg: Erase myZlonAdjstPos
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'クラス内Ｐ_2配列の傾き＆切片を計算
Private Sub CsubCalRegressionLineSlopeIntercept( _
            myXdouSlp As Double, myXdouIntrcpt As Double, _
            ByVal myXlonBgnNo As Long, ByVal myXlonEndNo As Long)
    myXdouSlp = Empty: myXdouIntrcpt = Empty
    
  Dim myXlonTmpXCnt As Long, myZdouTmpX() As Double
  Dim myXlonTmpYCnt As Long, myZdouTmpY() As Double

    If myXlonBgnNo = myXlonEndNo Then
        If myXlonBgnNo * myXlonEndNo = 1 Then
            myXlonBgnNo = myXlonBgnNo
            myXlonEndNo = myXlonEndNo + 1
        Else
            myXlonBgnNo = myXlonBgnNo - 1
            myXlonEndNo = myXlonEndNo
        End If
    End If
    
    Call Pfix1DArrayExtractNumeric(myXlonTmpXCnt, myZdouTmpX, _
            myZdouOrgX, myXlonBgnNo, myXlonEndNo, myXlonCalStep)
    If myXlonTmpXCnt <= 0 Then GoTo ExitPath
    Call Pfix1DArrayExtractNumeric(myXlonTmpYCnt, myZdouTmpY, _
            myZdouOrgY, myXlonBgnNo, myXlonEndNo, myXlonCalStep)
    If myXlonTmpYCnt <= 0 Then GoTo ExitPath
    
    Call PfixSlopeIntercept(myXdouSlp, myXdouIntrcpt, myZdouTmpX, myZdouTmpY)
    
    Erase myZdouTmpX: Erase myZdouTmpY
ExitPath:
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

'◆ClassProc名_データ配列の指定区間ごとで回帰直線間の差分データを取得する
Private Sub instCIntrvlRgrsnLnDev()
  Const coXlonCalDataRngOptn As Long = 1
    'coXlonCalRng = 1 : 全データ範囲で取得
    'coXlonCalRng = 2 : 指定データ範囲で取得
    'coXlonCalRng = 3 : 有効データ範囲で取得
  Const coXlonXCalRowCol As Long = 1
  Const coXlonYCalRowCol As Long = 5
  Const coXlonCalSmplQnt As Long = 11
  Const coXlonCalStep As Long = 1
  Const coXlonBgnRowCol As Long = 1
  Const coXlonEndRowCol As Long = 0
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
'    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
'    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
'  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
'  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
'//CIntrvlRgrsnLnDevクラスモジュールのインスタンス生成
  Dim myXinsIntrvlRgrsnLnDev As CIntrvlRgrsnLnDev
    Set myXinsIntrvlRgrsnLnDev = New CIntrvlRgrsnLnDev
  Dim i As Long, j As Long
    With myXinsIntrvlRgrsnLnDev
    '//クラス内変数への入力
        .letRowDrctn = True
        .letCalDataRngOptn = coXlonCalDataRngOptn
        If myXlonOrgRowCnt <= 0 Or myXlonOrgColCnt <= 0 Then GoTo JumpPath
        .letOrgRowCnt = myXlonOrgRowCnt
        .letOrgColCnt = myXlonOrgColCnt
        For j = LBound(myZdouOrgData, 2) To UBound(myZdouOrgData, 2)
            For i = LBound(myZdouOrgData, 1) To UBound(myZdouOrgData, 1)
                .letOrgData(i, j) = myZdouOrgData(i, j)
            Next i
        Next j
        .letXCalRowCol = coXlonXCalRowCol
        .letYCalRowCol = coXlonYCalRowCol
        .letCalSmplQnt = coXlonCalSmplQnt
        .letCalStep = coXlonCalStep
        .letBgnRowCol = coXlonBgnRowCol
        .letEndRowCol = coXlonEndRowCol
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonStrtRowCol = .getStrtRowCol
        myXlonLastRowCol = .getLastRowCol
        myXlonCalRowCnt = .getCalRowCnt
        myXlonCalColCnt = .getCalColCnt
        If myXlonCalRowCnt <= 0 Or myXlonCalColCnt <= 0 Then GoTo JumpPath
        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
        For j = 1 To myXlonCalColCnt
            For i = 1 To myXlonCalRowCnt
                myZdouCalData(i, j) = .getCalData(i, j)
            Next i
        Next j
    End With
JumpPath:
    Set myXinsIntrvlRgrsnLnDev = Nothing
    Call variablesOfCIntrvlRgrsnLnDev(myXlonCalRowCnt, myZdouCalData)    'Debug.Print
End Sub
Private Sub variablesOfCIntrvlRgrsnLnDev( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//CIntrvlRgrsnLnDevクラス内から出力した変数の内容確認
    Debug.Print "データ数: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim i As Long, j As Long
    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
        Next j
    Next i
End Sub

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
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
'    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
'    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
'  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
'  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
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
''    Call variablesOfCRgrsnLn(myXlonCalRowCnt, myZdouCalData)   'Debug.Print
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
