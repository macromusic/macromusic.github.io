VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHghPssFltr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsMatrix
'Includes PfixArrayPartColumnNumeric
'Includes PfixArrayPartRowNumeric
'Includes PfncdouHighPassFilter

Option Explicit
Option Base 1

'◆ClassProc名_データ配列をHighPassFilter処理する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CHghPssFltr"
  
'//PropertyGet_出力データ
  Private myXlonStrtRowCol As Long, myXlonLastRowCol As Long
  Private myXlonCalRowCnt As Long, myXlonCalColCnt As Long, _
            myZdouCalData() As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean, myXlonCalDataRngOptn As Long
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonCalRowCol As Long
  Private myXlonBgnRowCol As Long, myXlonEndRowCol As Long
  Private myXdouTs As Double, myXdouFc As Double
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, _
            myZdouOrgData() As Double
  Private myXlonOrgTmpCnt As Long, myZdouOrgTmp() As Double
  Private myXlonStrtNo As Long, myXlonLastNo As Long
  Private myXlonCalStrtNo As Long, myXlonCalLastNo As Long
  Private myXlonCalTmpCnt As Long, myZdouCalTmp() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgRowCnt = Empty: myXlonOrgColCnt = Empty: Erase myZdouOrgData
    myXlonOrgTmpCnt = Empty: Erase myZdouOrgTmp
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
'myZdouCalData(i, j) : 計算処理後データ配列
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

Public Property Let letCalRowCol(ByVal RHS As Long)
    myXlonCalRowCol = RHS
End Property

Public Property Let letBgnRowCol(ByVal RHS As Long)
    myXlonBgnRowCol = RHS
End Property

Public Property Let letEndRowCol(ByVal RHS As Long)
    myXlonEndRowCol = RHS
End Property

Public Property Let letTs(ByVal RHS As Double)
    myXdouTs = RHS
'myXdouTs[sec] : 元データ配列のサンプリング時間
End Property

Public Property Let letFc(ByVal RHS As Double)
    myXdouFc = RHS
'myXdouFc[Hz]  : HighPassFilterのカットオフ周波数
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
        If myXlonCalRowCol < LBound(myZdouOrgData, 2) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 2) Then GoTo ExitPath
        If myXlonBgnRowCol < LBound(myZdouOrgData, 1) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 1)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 1) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 1)
    Else
        If myXlonCalRowCol < LBound(myZdouOrgData, 1) _
            Or myXlonCalRowCol > UBound(myZdouOrgData, 1) Then GoTo ExitPath
        If myXlonBgnRowCol < LBound(myZdouOrgData, 2) Then _
            myXlonBgnRowCol = LBound(myZdouOrgData, 2)
        If myXlonEndRowCol < myXlonBgnRowCol _
                Or myXlonEndRowCol > UBound(myZdouOrgData, 2) Then _
            myXlonEndRowCol = UBound(myZdouOrgData, 2)
    End If
    
    If myXdouFc * myXdouTs <= 0 Then GoTo ExitPath
    
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

'//P:計算実行:HighPassFilter
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
    
'//計算が有効な最初と最後の位置を設定
    myXlonStrtNo = myXlonStrtNo + 1
    myXlonLastNo = myXlonLastNo
    If myXlonLastNo < myXlonStrtNo Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_計算実行:HighPassFilter
Private Sub prsCal()
    myXbisExitFlag = False
    
'//計算処理:HighPassFilter
  Dim i As Long, n As Long: n = 0
  Dim myXlonRetCnt As Long, myZdouRet() As Double
  Dim myXdouYnm1 As Double, myXdouXnm1 As Double, myXdouXn As Double

    myXlonRetCnt = UBound(myZdouOrgTmp) - LBound(myZdouOrgTmp) + 1
    ReDim myZdouRet(myXlonRetCnt) As Double
    
    For i = LBound(myZdouOrgTmp) To UBound(myZdouOrgTmp)
        n = n + 1
        If i = LBound(myZdouOrgTmp) Then
            myXdouYnm1 = myZdouOrgTmp(i)
            myXdouXnm1 = myZdouOrgTmp(i)
        Else
            myXdouYnm1 = myZdouRet(n - 1)
            myXdouXnm1 = myZdouOrgTmp(i - 1)
        End If
        myXdouXn = myZdouOrgTmp(i)
        myZdouRet(n) = PfncdouHighPassFilter( _
                        myXdouYnm1, myXdouXnm1, myXdouXn, myXdouTs, myXdouFc)
    Next i
    If n <> myXlonRetCnt Then GoTo ExitPath
    
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
    
'//データ抽出
    myXlonCalTmpCnt = myXlonCalLastNo - myXlonCalStrtNo + 1
    ReDim myZdouCalTmp(myXlonCalTmpCnt) As Double
    
    n = 0
    For i = myXlonCalStrtNo To myXlonCalLastNo
        n = n + 1
        myZdouCalTmp(n) = myZdouRet(i)
    Next i
    If n <> myXlonCalTmpCnt Then GoTo ExitPath
    
    Erase myZdouRet
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
'myZdouExtrctData(i) : 抽出データ配列
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

 '定型Ｆ_配列計算-HighPassFilter
Private Function PfncdouHighPassFilter( _
            ByVal myXdouYnm1 As Double, ByVal myXdouXnm1 As Double, _
            ByVal myXdouXn As Double, _
            ByVal myXdouTs As Double, ByVal myXdouFc As Double) As Double
'//計算式
    'Y[n]=1/(1+2*π*fc*Ts)*{Y[n-1]-X[n-1]+X[n]}
'myXdouYnm1 : 計算値の前回値
'myXdouXnm1 : 元データ前回値
'myXdouXn   : 元データ今回値
'myXdouTs[sec] : 元データ配列のサンプリング時間
'myXdouFc[Hz]  : HighPassFilterのカットオフ周波数
  Const myXdouPI As Double = 3.14159265358979
'  Dim myXdouPI As Double: myXdouPI = Application.WorksheetFunction.Pi()
    PfncdouHighPassFilter = Empty
  Dim myXdouTmpFc As Double
    myXdouTmpFc = myXdouFc * myXdouTs
    If myXdouTmpFc <= 0 Then Exit Function
  Dim myXdouRet As Double, myXdouA As Double
    myXdouA = 2 * myXdouPI * myXdouTmpFc
    myXdouA = 1 / (1 + myXdouA)
    myXdouRet = myXdouA * (myXdouYnm1 - myXdouXnm1 + myXdouXn)
    PfncdouHighPassFilter = myXdouRet
End Function

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_データ配列をHighPassFilter処理する
'Private Sub instCHghPssFltr()
'  Const coXlonCalDataRngOptn As Long = 1
'    'coXlonCalDataRngOptn = 1 : 全データ範囲で取得
'    'coXlonCalDataRngOptn = 2 : 指定データ範囲で取得
'    'coXlonCalDataRngOptn = 3 : 有効データ範囲で取得
'  Const coXlonCalRowCol As Long = 4
'  Const coXlonBgnRowCol As Long = 1
'  Const coXlonEndRowCol As Long = 0
'  Const coXdouTs As Double = 0.00001
'    'coXdouTs[sec] : 元データ配列のサンプリング時間
'  Const coXdouFc As Double = 500
'    'coXdouFc[Hz]  : HighPassFilterのカットオフ周波数
''  Dim myXlonOrgRowCnt As Long, myXlonOrgColCnt As Long, myZdouOrgData() As Double
''    myXlonOrgRowCnt = UBound(myZdouOrgData, 1)
''    myXlonOrgColCnt = UBound(myZdouOrgData, 2)
''  Dim myXlonStrtRowCol As Long, myXlonLastRowCol As Long
''  Dim myXlonCalRowCnt As Long, myXlonCalColCnt As Long, myZdouCalData() As Double
''//CHghPssFltrクラスモジュールのインスタンス生成
'  Dim myXinsHghPssFltr As CHghPssFltr: Set myXinsHghPssFltr = New CHghPssFltr
'  Dim i As Long, j As Long
'    With myXinsHghPssFltr
'    '//クラス内変数への入力
'        .letRowDrctn = True
'        .letCalDataRngOptn = coXlonCalDataRngOptn
'        .letOrgData = myZdouOrgData
'        .letCalRowCol = coXlonCalRowCol
'        .letBgnRowCol = coXlonBgnRowCol
'        .letEndRowCol = coXlonEndRowCol
'        .letTs = coXdouTs
'        .letFc = coXdouFc
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
'    Set myXinsHghPssFltr = Nothing
''    Call variablesOfCHghPssFltr(myXlonCalRowCnt, myZdouCalData)  'Debug.Print
'End Sub
'Private Sub variablesOfCHghPssFltr( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CHghPssFltrクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
