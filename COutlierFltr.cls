VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COutlierFltr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes PfixArrayPartColumnNumeric
'Includes PfixArrayPartRowNumeric
'Includes Pfix1DArrayExtractNumeric
'Includes PfixTrapezoidWaveLong
'Includes PfixSlopeIntercept
'Includes PfncdouStDevS
'Includes PfncdouArithmeticMean
'Includes incPfncdouMedian
'Includes PfixQuickSortOfNumeric1DArray

Option Explicit
Option Base 1

'◆ClassProc名_データ配列から標準偏差外れのデータを除去する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "COutlierFltr"
  
'//モジュール内定数
  Private Const coXlonMdnFltrQnt As Long = 3
  
'//PropertyGet_出力データ
  Private myXlonStrtRowCol As Long, myXlonLastRowCol As Long
  Private myXlonCalRowCnt As Long, myXlonCalColCnt As Long, _
            myZdouCalData() As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean, myXlonCalRng As Long, myXlonCalOptn As Long
  Private myXlonCalDataRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonCalRowCol As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  Private myXlonRgrsnSmplQnt As Long, myXlonRgrsnStep As Long
  Private myXlonMeanSmplQnt As Long, myXlonMeanStep As Long
  Private myXlonMedianSmplQnt As Long, myXlonMedianStep As Long
  Private myXdouSigmaLimit As Double
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double
  
  Private myXlonXdataCnt As Long, myZdouXdata() As Double
  Private myXlonOrgYCnt As Long, myZdouOrgY() As Double
  Private myXlonStrtNo As Long, myXlonLastNo As Long
  Private myXlonPreOfst As Long, myXlonPstOfst As Long
  Private myXlonCalStrtNo As Long, myXlonCalLastNo As Long
  
  Private myXlonRgrsnLnDevCnt As Long, myZdouRgrsnLnDev() As Double
  Private myXdouStDev As Double
  Private myXlonMeanCnt As Long, myZdouMean() As Double
  Private myXlonMedianCnt As Long, myZdouMedian() As Double
  Private myXlonFltrDataCnt As Long, myZdouFltrData() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgRowCnt = Empty: myXlonOrgColCnt = Empty: Erase myZdouOrgData
    
    myXlonXdataCnt = Empty: Erase myZdouXdata
    myXlonOrgYCnt = Empty: Erase myZdouOrgY
    myXlonStrtNo = Empty: myXlonLastNo = Empty
    myXlonPreOfst = Empty: myXlonPstOfst = Empty
    myXlonCalStrtNo = Empty: myXlonCalLastNo = Empty

    myXlonRgrsnLnDevCnt = Empty: Erase myZdouRgrsnLnDev
    myXdouStDev = Empty
    myXlonMeanCnt = Empty: Erase myZdouMean
    myXlonMedianCnt = Empty: Erase myZdouMedian
    myXlonFltrDataCnt = Empty: Erase myZdouFltrData
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

Public Property Let letCalRng(ByVal RHS As Long)
    myXlonCalRng = RHS
'myXlonCalRng = -1 : 今回値より前のデータで計算
'myXlonCalRng = 0  : 今回値の前後のデータで計算
'myXlonCalRng = 1  : 今回値より後のデータで計算
End Property

Public Property Let letCalOptn(ByVal RHS As Long)
    myXlonCalOptn = RHS
'myXlonCalOptn = 1 : 平均値を基準に標準偏差外れのデータを除去
'myXlonCalOptn = 2 : 中央値を基準に標準偏差外れのデータを除去
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

Public Property Let letBgnRowCol(ByVal RHS As Long)
    myXlonBgnRowCol = RHS
End Property

Public Property Let letEndRowCol(ByVal RHS As Long)
    myXlonEndRowCol = RHS
End Property

Public Property Let letRgrsnSmplQnt(ByVal RHS As Long)
    myXlonRgrsnSmplQnt = RHS
'myXlonRgrsnSmplQnt : 回帰直線処理するデータ数
End Property

Public Property Let letRgrsnStep(ByVal RHS As Long)
    myXlonRgrsnStep = RHS
End Property

Public Property Let letMeanSmplQnt(ByVal RHS As Long)
    myXlonMeanSmplQnt = RHS
'myXlonMeanSmplQnt : 平均値処理するデータ数
End Property

Public Property Let letMeanStep(ByVal RHS As Long)
    myXlonMeanStep = RHS
End Property

Public Property Let letMedianSmplQnt(ByVal RHS As Long)
    myXlonMedianSmplQnt = RHS
'myXlonMedianSmplQnt : メジアン処理するデータ数
End Property

Public Property Let letMedianStep(ByVal RHS As Long)
    myXlonMedianStep = RHS
End Property

Public Property Let letSigmaLimit(ByVal RHS As Double)
    myXdouSigmaLimit = RHS
'myXdouSigmaLimit : 除外する外れ値の範囲を標準偏差σで指定
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXlonCalRng < -1 Or myXlonCalRng > 1 Then myXlonCalRng = 0
    If myXlonCalOptn < 1 Or myXlonCalOptn > 2 Then myXlonCalOptn = 1
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
        If myXlonCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonRgrsnSmplQnt <= 0 _
            Or myXlonRgrsnSmplQnt > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonRgrsnStep <= 0 Or myXlonRgrsnStep > UBound(myZdouOrgData, 1) Then _
            myXlonRgrsnStep = 1
        If myXlonMeanSmplQnt <= 0 _
            Or myXlonMeanSmplQnt > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonMeanStep <= 0 Or myXlonMeanStep > UBound(myZdouOrgData, 1) Then _
            myXlonMeanStep = 1
        If myXlonMedianSmplQnt <= 0 _
            Or myXlonMedianSmplQnt > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonMedianStep <= 0 Or myXlonMedianStep > UBound(myZdouOrgData, 1) Then _
            myXlonMedianStep = 1
    Else
        If myXlonBgnRowCol < LBound(myZdouOrgData, 2) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 2)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 2) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 2)
        If myXlonCalRowCol < LBound(myZdouOrgData, 1) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonRgrsnSmplQnt <= 0 _
            Or myXlonRgrsnSmplQnt > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonRgrsnStep <= 0 Or myXlonRgrsnStep > UBound(myZdouOrgData, 2) Then _
            myXlonRgrsnStep = 1
        If myXlonMeanSmplQnt <= 0 _
            Or myXlonMeanSmplQnt > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonMeanStep <= 0 Or myXlonMeanStep > UBound(myZdouOrgData, 2) Then _
            myXlonMeanStep = 1
        If myXlonMedianSmplQnt <= 0 _
            Or myXlonMedianSmplQnt > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonMedianStep <= 0 Or myXlonMedianStep > UBound(myZdouOrgData, 2) Then _
            myXlonMedianStep = 1
    End If
    
    If myXdouSigmaLimit <= 0 Then myXdouSigmaLimit = 3
    
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
    
'//P:計算実行:Y軸データと回帰直線との差分
    Call prsCalDeviationFromRegressionLine
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:計算実行:標準偏差
    Call prsCalStandardDeviation
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    If myXlonCalOptn = 1 Then
    '//P:計算実行:平均値
        Call prsCalMean
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
        
    '//P:平均値基準で標準偏差外れ値をフィルタ処理
        Call prsProcStandardDeviationOutliersFilterWithMean
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "6"     'PassFlag
        
    Else
    '//P:計算実行:中央値
        Call prsCalMedian
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "7"     'PassFlag
        
    '//P:中央値基準で標準偏差外れ値をフィルタ処理
        Call prsProcStandardDeviationOutliersFilterWithMedian
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "8"     'PassFlag
        
    End If
    
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
        Call PfixArrayPartColumnNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonCalRowCol, myXlonBgnNo, myXlonEndNo)
    Else
        myXlonBgnNo = LBound(myZdouOrgData, 2)
        myXlonEndNo = UBound(myZdouOrgData, 2)
        Call PfixArrayPartRowNumeric(myXlonOrgYCnt, myZdouOrgY, _
                myZdouOrgData, myXlonCalRowCol, myXlonBgnNo, myXlonEndNo)
    End If
    If myXlonOrgYCnt <= 0 Then GoTo ExitPath
    
  Dim i As Long
    myXlonXdataCnt = myXlonOrgYCnt
    ReDim myZdouXdata(myXlonOrgYCnt) As Double
    For i = LBound(myZdouOrgY) To UBound(myZdouOrgY)
        myZdouXdata(i) = i
    Next i
    
    myXlonStrtNo = myXlonBgnNo
    myXlonLastNo = myXlonEndNo
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables( _
            ByVal myXlonCalRng As Long, _
            ByVal myXlonCalSmplQnt As Long, ByVal myXlonCalStep As Long)
    myXbisExitFlag = False
    myXlonPreOfst = Empty: myXlonPstOfst = Empty
    
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

'PrcsP_計算実行:Y軸データと回帰直線との差分
Private Sub prsCalDeviationFromRegressionLine()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Call setControlVariables(myXlonCalRng, myXlonRgrsnSmplQnt, myXlonRgrsnStep)
    If myXbisExitFlag = True Then GoTo ExitPath
  Dim myXlonCalStrtNo As Long, myXlonCalLastNo As Long
    myXlonCalStrtNo = LBound(myZdouOrgY) + myXlonPreOfst
    myXlonCalLastNo = UBound(myZdouOrgY) - myXlonPstOfst
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgYCnt, myXlonCalStrtNo, myXlonCalLastNo, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
'//計算処理:Y軸データと回帰直線との差分
  Dim i As Long, myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonXTmpCnt As Long, myZdouXTmp() As Double
  Dim myXlonYTmpCnt As Long, myZdouYTmp() As Double
  Dim myXdouSlp As Double, myXdouIntrcpt As Double, myXdouRgrsn As Double
    
    myXlonRgrsnLnDevCnt = myXlonOrgYCnt
    ReDim myZdouRgrsnLnDev(myXlonRgrsnLnDevCnt) As Double

    For i = LBound(myZdouOrgY) To UBound(myZdouOrgY)
        myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
        myXlonEndNo = i + myXlonPstOfst + myZlonAdjstVal(i)
        Call Pfix1DArrayExtractNumeric(myXlonXTmpCnt, myZdouXTmp, _
                myZdouXdata, myXlonBgnNo, myXlonEndNo, myXlonRgrsnStep)
        If myXlonXTmpCnt <= 0 Then GoTo ExitPath
        Call Pfix1DArrayExtractNumeric(myXlonYTmpCnt, myZdouYTmp, _
                myZdouOrgY, myXlonBgnNo, myXlonEndNo, myXlonRgrsnStep)
        If myXlonYTmpCnt <= 0 Then GoTo ExitPath
        Call PfixSlopeIntercept(myXdouSlp, myXdouIntrcpt, myZdouXTmp, myZdouYTmp)
        myXdouRgrsn = myXdouSlp * myZdouXdata(i) + myXdouIntrcpt
        myZdouRgrsnLnDev(i) = myZdouOrgY(i) - myXdouRgrsn
    Next i
    
    Erase myZlonAdjstVal: Erase myZdouXTmp: Erase myZdouYTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:標準偏差
Private Sub prsCalStandardDeviation()
    myXbisExitFlag = False
    
'//計算処理:不偏標本標準偏差
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
            myZdouRgrsnLnDev, myXlonStrtNo, myXlonLastNo)
    If myXlonTmpCnt <= 0 Then GoTo ExitPath
    myXdouStDev = PfncdouStDevS(myZdouTmp)
    
    Erase myZdouTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:平均値
Private Sub prsCalMean()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Call setControlVariables(myXlonCalRng, myXlonMeanSmplQnt, myXlonMeanStep)
    If myXbisExitFlag = True Then GoTo ExitPath
  Dim myXlonCalStrt As Long, myXlonCalLast As Long
    myXlonCalStrt = LBound(myZdouOrgY) + myXlonPreOfst
    myXlonCalLast = UBound(myZdouOrgY) - myXlonPstOfst
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgYCnt, myXlonCalStrt, myXlonCalLast, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
'//計算処理:平均値
  Dim i As Long, myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonMeanCnt = myXlonOrgYCnt
    ReDim myZdouMean(myXlonMeanCnt) As Double

    For i = LBound(myZdouOrgY) To UBound(myZdouOrgY)
        myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
        myXlonEndNo = i + myXlonPstOfst + myZlonAdjstVal(i)
        Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                myZdouRgrsnLnDev, myXlonBgnNo, myXlonEndNo, myXlonMeanStep)
        If myXlonTmpCnt <= 0 Then GoTo ExitPath
        myZdouMean(i) = PfncdouArithmeticMean(myZdouTmp)
    Next i
    
    Erase myZlonAdjstVal: Erase myZdouTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:中央値
Private Sub prsCalMedian()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Call setControlVariables(myXlonCalRng, myXlonMedianSmplQnt, myXlonMedianStep)
    If myXbisExitFlag = True Then GoTo ExitPath
  Dim myXlonCalStrt As Long, myXlonCalLast As Long
    myXlonCalStrt = LBound(myZdouOrgY) + myXlonPreOfst
    myXlonCalLast = UBound(myZdouOrgY) - myXlonPstOfst
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgYCnt, myXlonCalStrt, myXlonCalLast, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
'//計算処理:中央値
  Dim i As Long, myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonMedianCnt = myXlonOrgYCnt
    ReDim myZdouMedian(myXlonMedianCnt) As Double

    For i = LBound(myZdouOrgY) To UBound(myZdouOrgY)
        myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
        myXlonEndNo = i + myXlonPstOfst + myZlonAdjstVal(i)
        Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                myZdouRgrsnLnDev, myXlonBgnNo, myXlonEndNo, myXlonMeanStep)
        If myXlonTmpCnt <= 0 Then GoTo ExitPath
        myZdouMedian(i) = incPfncdouMedian(myZdouTmp)
    Next i
    
    Erase myZlonAdjstVal: Erase myZdouTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_平均値基準で標準偏差外れ値をフィルタ処理
Private Sub prsProcStandardDeviationOutliersFilterWithMean()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Call setControlVariables(0, coXlonMdnFltrQnt, 1)
    If myXbisExitFlag = True Then GoTo ExitPath
    myXlonCalStrtNo = LBound(myZdouOrgY) + myXlonPreOfst
    myXlonCalLastNo = UBound(myZdouOrgY) - myXlonPstOfst
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgYCnt, myXlonCalStrtNo, myXlonCalLastNo, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = LBound(myZdouOrgY)
            myXlonCalLastNo = UBound(myZdouOrgY)
        Case 2
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
        Case Else
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
    End Select
    
'//標準偏差外れ値のフィルタ処理
  Dim i As Long, n As Long: n = 0
  Dim myXdouCalVal As Double, myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonFltrDataCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouFltrData(myXlonFltrDataCnt) As Double

    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        
    '//計算処理:｜回帰直線との差分－平均値｜／標準偏差
        myXdouCalVal = Abs(myZdouRgrsnLnDev(i) - myZdouMean(i)) / myXdouStDev
        
    '//計算処理:標準偏差のX倍より大きいか判断
        If myXdouCalVal <= myXdouSigmaLimit Then
        '//元データを取得
            myZdouFltrData(n) = myZdouOrgY(i)
            
        Else
        '//元データの前後3データ分の中央値を取得
            myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
            myXlonEndNo = i + myXlonPreOfst + myZlonAdjstVal(i)
            Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                    myZdouOrgY, myXlonBgnNo, myXlonEndNo)
            If myXlonTmpCnt <> coXlonMdnFltrQnt Then GoTo ExitPath
            myZdouFltrData(n) = incPfncdouMedian(myZdouTmp)
            
        End If
        
    Next i
    
    Erase myZlonAdjstVal: Erase myZdouTmp
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_中央値基準で標準偏差外れ値をフィルタ処理
Private Sub prsProcStandardDeviationOutliersFilterWithMedian()
    myXbisExitFlag = False
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Call setControlVariables(0, coXlonMdnFltrQnt, 1)
    If myXbisExitFlag = True Then GoTo ExitPath
    myXlonCalStrtNo = LBound(myZdouOrgY) + myXlonPreOfst
    myXlonCalLastNo = UBound(myZdouOrgY) - myXlonPstOfst
    
'//台形波形を生成
  Dim myXlonAdjstValCnt As Long, myZlonAdjstVal() As Long
    Call PfixTrapezoidWaveLong(myXlonAdjstValCnt, myZlonAdjstVal, _
            myXlonOrgYCnt, myXlonCalStrtNo, myXlonCalLastNo, -1, -1, 0)
    If myXlonAdjstValCnt <> myXlonOrgYCnt Then GoTo ExitPath
    
'//計算する最初のデータ位置と最後のデータ位置を設定
    Select Case myXlonCalDataRngOptn
        Case 1
            myXlonCalStrtNo = LBound(myZdouOrgY)
            myXlonCalLastNo = UBound(myZdouOrgY)
        Case 2
            myXlonCalStrtNo = myXlonBgnRowCol
            myXlonCalLastNo = myXlonEndRowCol
        Case Else
            myXlonCalStrtNo = myXlonStrtNo
            myXlonCalLastNo = myXlonLastNo
    End Select
    
'//標準偏差外れ値のフィルタ処理
  Dim i As Long, n As Long: n = 0
  Dim myXdouCalVal As Double, myXlonBgnNo As Long, myXlonEndNo As Long
  Dim myXlonTmpCnt As Long, myZdouTmp() As Double
    
    myXlonFltrDataCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouFltrData(myXlonFltrDataCnt) As Double
    
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
    
    '//計算処理:｜回帰直線との差分－平均値｜／標準偏差
        myXdouCalVal = Abs(myZdouRgrsnLnDev(i) - myZdouMedian(i)) / myXdouStDev
        
    '//計算処理:標準偏差のX倍より大きいか判断
        If myXdouCalVal <= myXdouSigmaLimit Then
        '//元データを取得
            myZdouFltrData(n) = myZdouOrgY(i)
            
        Else
        '//元データの前後3データ分の中央値を取得
            myXlonBgnNo = i - myXlonPreOfst + myZlonAdjstVal(i)
            myXlonEndNo = i + myXlonPreOfst + myZlonAdjstVal(i)
            Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
                    myZdouOrgY, myXlonBgnNo, myXlonEndNo)
            If myXlonTmpCnt <> coXlonMdnFltrQnt Then GoTo ExitPath
            myZdouFltrData(n) = incPfncdouMedian(myZdouTmp)
            
        End If
        
    Next i
    
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
        myXlonCalRowCnt = myXlonFltrDataCnt
        myXlonCalColCnt = 1
        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
        For i = LBound(myZdouFltrData) To UBound(myZdouFltrData)
            myZdouCalData(i, 1) = myZdouFltrData(i)
        Next i
        
    Else
        myXlonCalRowCnt = 1
        myXlonCalColCnt = myXlonFltrDataCnt
        ReDim myZdouCalData(myXlonCalRowCnt, myXlonCalColCnt) As Double
        For j = LBound(myZdouFltrData) To UBound(myZdouFltrData)
            myZdouCalData(1, j) = myZdouFltrData(j)
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

 '定型Ｆ_配列計算-不偏標本標準偏差
Private Function PfncdouStDevS(ByRef myZdouOrgData() As Double) As Double
'myZdouOrgData(i) : 元データ配列
    PfncdouStDevS = Empty
  Dim myXdouSumSq As Double: myXdouSumSq = 0
  Dim myXdouSum As Double: myXdouSum = 0
  Dim i As Long, n As Long: n = 0
    For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
        n = n + 1
        myXdouSumSq = myXdouSumSq + (myZdouOrgData(i)) ^ 2
        myXdouSum = myXdouSum + myZdouOrgData(i)
    Next i
    If n <= 1 Then Exit Function
    PfncdouStDevS = ((myXdouSumSq - (myXdouSum ^ 2) / n) / (n - 1)) ^ 0.5
End Function

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

 '定型Ｆ_配列計算-中央値
Private Function incPfncdouMedian(ByRef myZdouOrgData() As Double) As Double
'Includes PfixQuickSortOfNumeric1DArray
'myZdouOrgData(i) : 元データ配列
    incPfncdouMedian = Empty
  Dim myXlonLbnd As Long: myXlonLbnd = LBound(myZdouOrgData)
  Dim myXlonUbnd As Long: myXlonUbnd = UBound(myZdouOrgData)
  Dim myXlonCnt As Long: myXlonCnt = myXlonUbnd - myXlonLbnd + 1
  Dim myZdouTmp() As Double, i As Long
    ReDim myZdouTmp(myXlonUbnd) As Double
    For i = myXlonLbnd To myXlonUbnd
        myZdouTmp(i) = myZdouOrgData(i)
    Next i
    Call PfixQuickSortOfNumeric1DArray(myZdouTmp, myXlonLbnd, myXlonUbnd, 1)
    If myXlonCnt Mod 2 = 1 Then
        incPfncdouMedian = myZdouTmp(Int((myXlonLbnd + myXlonUbnd) / 2))
    Else
      Dim m As Long, n As Long, myXdouM1 As Double, myXdouM2 As Double
        m = Int((myXlonLbnd + myXlonUbnd) / 2): n = m + 1
        myXdouM1 = myZdouTmp(m): myXdouM2 = myZdouTmp(n)
        incPfncdouMedian = (myXdouM1 + myXdouM2) / 2
    End If
    Erase myZdouTmp
End Function

 '定型Ｐ_1次元の数値配列のクイックソート
Private Sub PfixQuickSortOfNumeric1DArray( _
            myZdouSrtData() As Double, _
            ByVal myXlonLbnd As Long, ByVal myXlonUbnd As Long, _
            Optional ByVal coXlonOrder As Long = 1)
'myZdouSrtData(i) : 数値情報
'coXlonOrder = 1    : 昇順 (Ascending Order)
'coXlonOrder = -1   : 降順 (Descending Order)
    If myXlonUbnd <= myXlonLbnd Then Exit Sub
  Dim i As Long, j As Long
  Dim myXdouMed As Double, myXdouTmp As Double
    myXdouMed = myZdouSrtData(Int((myXlonLbnd + myXlonUbnd) / 2))
    i = myXlonLbnd: j = myXlonUbnd
    Do
        Do While coXlonOrder * (myZdouSrtData(i) - myXdouMed) < 0
            i = i + 1
        Loop
        Do While coXlonOrder * (myZdouSrtData(j) - myXdouMed) > 0
            j = j - 1
        Loop
        If i >= j Then Exit Do
        myXdouTmp = myZdouSrtData(i)
        myZdouSrtData(i) = myZdouSrtData(j)
        myZdouSrtData(j) = myXdouTmp
        i = i + 1: j = j - 1
    Loop
    If myXlonLbnd < i - 1 Then _
        Call PfixQuickSortOfNumeric1DArray(myZdouSrtData, myXlonLbnd, i - 1, coXlonOrder)
    If myXlonUbnd > j + 1 Then _
        Call PfixQuickSortOfNumeric1DArray(myZdouSrtData, j + 1, myXlonUbnd, coXlonOrder)
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_データ配列から標準偏差外れのデータを除去する
'Private Sub instCOutlierFltr()
'  Const coXlonCalOptn As Long = 1
'    'coXlonCalOptn = 1 : 平均値を基準に標準偏差外れのデータを除去
'    'coXlonCalOptn = 2 : 中央値を基準に標準偏差外れのデータを除去
'  Const coXlonCalRng As Long = 0
'    'coXlonCalRng = -1 : 今回値より前のデータで平均化
'    'coXlonCalRng = 0  : 今回値の前後のデータで平均化
'    'coXlonCalRng = 1  : 今回値より後のデータで平均化
'  Const coXlonCalDataRngOptn As Long = 1
'    'coXlonCalDataRngOptn = 1 : 全データ範囲で取得
'    'coXlonCalDataRngOptn = 2 : 指定データ範囲で取得
'    'coXlonCalDataRngOptn = 3 : 有効データ範囲で取得
'  Const coXlonCalRowCol As Long = 5
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
'  Const coXlonRgrsnSmplQnt As Long = 11
'    'coXlonRgrsnSmplQnt : 回帰直線処理するデータ数
'  Const coXlonRgrsnStep As Long = 1
'  Const coXlonMeanSmplQnt As Long = 11
'    'coXlonMeanSmplQnt : 平均値処理するデータ数
'  Const coXlonMeanStep As Long = 1
'  Const coXlonMedianSmplQnt As Long = 11
'    'coXlonMedianSmplQnt : メジアン処理するデータ数
'  Const coXlonMedianStep As Long = 1
'  Const coXdouSigmaLimit As Long = 3
'    'coXdouSigmaLimit : 除外する外れ値の範囲を標準偏差σで指定
'  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
'    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
'    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
'  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
'  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
''//COutlierFltrクラスモジュールのインスタンス生成
'  Dim myXinsOutlierFltr As COutlierFltr: Set myXinsOutlierFltr = New COutlierFltr
'  Dim i As Long, j As Long
'    With myXinsOutlierFltr
'    '//クラス内変数への入力
'        .letRowDrctn = True
'        .letCalOptn = coXlonCalOptn
'        .letCalRng = coXlonCalRng
'        .letCalDataRngOptn = coXlonCalDataRngOptn
'        .letOrgData = myZdouOrgData
'        .letCalRowCol = myXlonCalRowCol
'        .letBgnRowCol = coXlonBgnRowCol
'        .letEndRowCol = coXlonEndRowCol
'        .letRgrsnSmplQnt = myXlonRgrsnSmplQnt
'        .letRgrsnStep = myXlonRgrsnStep
'        .letMeanSmplQnt = myXlonMeanSmplQnt
'        .letMeanStep = myXlonMeanStep
'        .letMedianSmplQnt = myXlonMedianSmplQnt
'        .letMedianStep = myXlonMedianStep
'        .letSigmaLimit = myXdouSigmaLimit
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
'    Set myXinsOutlierFltr = Nothing
''    Call variablesOfCOutlierFltr(myXlonCalRowCnt, myZdouCalData)    'Debug.Print
'End Sub
'Private Sub variablesOfCOutlierFltr( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//COutlierFltrクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
