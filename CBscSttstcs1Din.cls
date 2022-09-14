VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBscSttstcs1Din"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisIsNumeric1DArray
'Includes Pfix1DArrayExtractNumeric
'Includes PfncdouSum
'Includes incPfixMaxMinMedian
'Includes PfixQuickSortOfNumeric1DArray
'Includes PfncdouArithmeticMean
'Includes PfncdouStDevP
'Includes PfncdouStDevS
'Includes PfnclonMatchNumeric

Option Explicit
Option Base 1

'◆ClassProc名_1次元の数値配列内のデータから統計量を計算する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CBscSttstcs1Din"
  
'//PropertyGet_出力データ
  Private myXlonDataCnt As Long, myXdouSum As Double
  Private myXdouMax As Double, myXlonMaxDataNo As Long
  Private myXdouMin As Double, myXlonMinDataNo As Long
  Private myXdouMaxAbs As Double, myXlonMaxAbsDataNo As Long
  Private myXdouMedian As Double, myXdouArthmtcMean As Double
  Private myXdouStDev As Double
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean
  Private myXstrOutOptn As String
  
'//PropertySetLet_入力データ
  Private myZdouOrgDataIN() As Double
  Private myXlonBgn As Long, myXlonEnd As Long, myXlonCalStep As Long
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgCnt As Long, myZdouOrgData() As Double
  Private myXlonTmpCnt As Long, myZdouTmp() As Double

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgCnt = Empty: Erase myZdouOrgData
    myXlonTmpCnt = Empty: Erase myZdouTmp
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getDataCnt() As Long
    getDataCnt = myXlonDataCnt
End Property

Public Property Get getSum() As Double
    getSum = myXdouSum
End Property

Public Property Get getMax() As Double
    getMax = myXdouMax
End Property

Public Property Get getMaxDataNo() As Long
    getMaxDataNo = myXlonMaxDataNo
End Property

Public Property Get getMin() As Double
    getMin = myXdouMin
End Property

Public Property Get getMinDataNo() As Long
    getMinDataNo = myXlonMinDataNo
End Property

Public Property Get getMaxAbs() As Double
    getMaxAbs = myXdouMaxAbs
End Property

Public Property Get getMaxAbsDataNo() As Long
    getMaxAbsDataNo = myXlonMaxAbsDataNo
End Property

Public Property Get getMedian() As Double
    getMedian = myXdouMedian
End Property

Public Property Get getArthmtcMean() As Double
    getArthmtcMean = myXdouArthmtcMean
End Property

Public Property Get getStDev() As Double
    getStDev = myXdouStDev
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonDataCnt = Empty: myXdouSum = Empty
    myXdouMax = Empty: myXlonMaxDataNo = Empty
    myXdouMin = Empty: myXlonMinDataNo = Empty
    myXdouMaxAbs = Empty: myXlonMaxAbsDataNo = Empty
    myXdouMedian = Empty: myXdouArthmtcMean = Empty
    myXdouStDev = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letOrgData(ByRef RHS() As Double)
'myZdouOrgData(i) : 元データ配列
    myZdouOrgDataIN = RHS
End Property

Public Property Let letBgn(ByVal RHS As Long)
    myXlonBgn = RHS
End Property

Public Property Let letEnd(ByVal RHS As Long)
    myXlonEnd = RHS
End Property

Public Property Let letCalStep(ByVal RHS As Long)
    myXlonCalStep = RHS
End Property

Public Property Let letOutOptn(ByVal RHS As String)
    myXstrOutOptn = RHS
'myXstrOutOptn = "10000000000" : N
'myXstrOutOptn = "01000000000" : Sum
'myXstrOutOptn = "00100000000" : Max
'myXstrOutOptn = "00010000000" : MaxDataNo
'myXstrOutOptn = "00001000000" : Min
'myXstrOutOptn = "00000100000" : MinDataNo
'myXstrOutOptn = "00000010000" : MaxAbs
'myXstrOutOptn = "00000001000" : MaxAbsDataNo
'myXstrOutOptn = "00000000100" : Median
'myXstrOutOptn = "00000000010" : ArithmeticMean
'myXstrOutOptn = "00000000001" : StDev
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If PfncbisIsNumeric1DArray(myZdouOrgDataIN) = False Then GoTo ExitPath
    
  Dim L As Long, i As Long
    L = LBound(myZdouOrgDataIN)
    myXlonOrgCnt = UBound(myZdouOrgDataIN, 1) - L + 1
    ReDim myZdouOrgData(myXlonOrgCnt) As Double
    For i = LBound(myZdouOrgDataIN) To UBound(myZdouOrgDataIN)
        myZdouOrgData(i + (1 - L)) = myZdouOrgDataIN(i)
    Next i
    
    If myXlonBgn < LBound(myZdouOrgData) Then myXlonBgn = LBound(myZdouOrgData)
    If myXlonEnd < myXlonBgn Or myXlonEnd > UBound(myZdouOrgData) Then _
        myXlonEnd = UBound(myZdouOrgData)
    If myXlonCalStep <= 0 Or myXlonCalStep > UBound(myZdouOrgData) Then _
        myXlonCalStep = 1
    
    If Len(myXstrOutOptn) <> 11 Then GoTo ExitPath
    
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
    
'//S:計算するデータ配列を取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:統計量を計算
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'SnsP_計算するデータ配列を取得する
Private Sub snsProc()
    myXbisExitFlag = False
    
'//1次元の数値配列から一部を1次元配列で取得
    Call Pfix1DArrayExtractNumeric(myXlonTmpCnt, myZdouTmp, _
            myZdouOrgData, myXlonBgn, myXlonEnd, myXlonCalStep)
    If myXlonTmpCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_統計量を計算する
Private Sub prsProc()
    myXbisExitFlag = False
    
'//データ数
    If Mid(myXstrOutOptn, 1, 1) = "1" Then _
        myXlonDataCnt = UBound(myZdouTmp) - LBound(myZdouTmp) + 1
'    If Mid(myXstrOutOptn, 1, 1) = "1" Then _
'        myXlonDataCnt = Application.WorksheetFunction.Count(myZdouTmp)
    
'//配列計算:合計
    If Mid(myXstrOutOptn, 2, 1) = "1" Then _
        myXdouSum = PfncdouSum(myZdouTmp)
'    If Mid(myXstrOutOptn, 2, 1) = "1" Then _
'        myXdouSum = Application.WorksheetFunction.Sum(myZdouTmp)
    
'//配列計算:最大＆最小＆中央値
    If Mid(myXstrOutOptn, 3, 1) = "1" Or _
            Mid(myXstrOutOptn, 5, 1) = "1" Or _
            Mid(myXstrOutOptn, 9, 1) = "1" Then _
        Call incPfixMaxMinMedian(myXdouMax, myXdouMin, myXdouMedian, myZdouTmp)
'    If Mid(myXstrOutOptn, 3, 1) = "1" Then _
'        myXdouMax = Application.WorksheetFunction.Max(myZdouTmp)
'    If Mid(myXstrOutOptn, 5, 1) = "1" Then _
'        myXdouMin = Application.WorksheetFunction.Min(myZdouTmp)
'    If Mid(myXstrOutOptn, 9, 1) = "1" Then _
'        myXdouMedian = Application.WorksheetFunction.Median(myZdouTmp)
    
'//最大＆最小となるデータNo
    If Mid(myXstrOutOptn, 4, 1) = "1" Then _
        myXlonMaxDataNo = PfnclonMatchNumeric(myXdouMax, myZdouTmp, False)
    If Mid(myXstrOutOptn, 6, 1) = "1" Then _
        myXlonMinDataNo = PfnclonMatchNumeric(myXdouMin, myZdouTmp, False)
    
'//絶対値で最大となるデータ＆データNo
    If Mid(myXstrOutOptn, 7, 1) = "1" Then
        If Abs(myXdouMax) >= Abs(myXdouMin) Then
            myXdouMaxAbs = myXdouMax
        Else
            myXdouMaxAbs = myXdouMin
        End If
    End If
    If Mid(myXstrOutOptn, 8, 1) = "1" Then _
        myXlonMaxAbsDataNo = PfnclonMatchNumeric(myXdouMaxAbs, myZdouTmp, False)
    
'//配列計算:相加平均
    If Mid(myXstrOutOptn, 10, 1) = "1" Then _
        myXdouArthmtcMean = PfncdouArithmeticMean(myZdouTmp)
'    If Mid(myXstrOutOptn, 10, 1) = "1" Then _
'        myXdouArthmtcMean = Application.WorksheetFunction.Average(myZdouTmp)
    
'//配列計算:不偏標本標準偏差
    If Mid(myXstrOutOptn, 11, 1) = "1" Then _
        myXdouStDev = PfncdouStDevS(myZdouTmp)
'    If Mid(myXstrOutOptn, 11, 1) = "1" Then _
'        myXdouStDev = Application.WorksheetFunction.StDev(myZdouTmp)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

 '定型Ｆ_指定変数が1次元の数値配列かどうかをチェックする
Private Function PfncbisIsNumeric1DArray(ByRef myZvarAry As Variant) As Boolean
    PfncbisIsNumeric1DArray = False
  Dim myXdouTmp As Double, L As Long
    On Error GoTo ExitPath
    L = LBound(myZvarAry): myXdouTmp = myZvarAry(L)
    On Error GoTo 0
    PfncbisIsNumeric1DArray = True
ExitPath:
End Function

 '定型Ｐ_1次元の数値配列から一部を取得する
Private Sub Pfix1DArrayExtractNumeric( _
            myXlonExtrctCnt As Long, myZdouExtrctData() As Double, _
            ByRef myZdouOrgData() As Double, _
            ByVal myXlonBgn As Long, ByVal myXlonEnd As Long, _
            Optional ByVal coXlonStep As Long = 1)
'myZdouExtrctData(i) : 取得データ配列
'myZdouOrgData(i) : 元データ配列
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

 '定型Ｆ_配列計算-合計
Private Function PfncdouSum(ByRef myZdouOrgData() As Double) As Double
'myZdouOrgData(i) : 元データ配列
    PfncdouSum = Empty
  Dim i As Long, myXdouSum As Double: myXdouSum = 0
    For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
        myXdouSum = myXdouSum + myZdouOrgData(i)
    Next i
    PfncdouSum = myXdouSum
End Function

 '定型Ｐ_配列計算-最大＆最小＆中央値
Private Sub incPfixMaxMinMedian( _
            myXdouMax As Double, myXdouMin As Double, myXdouMedian As Double, _
            ByRef myZdouOrgData() As Double)
'Includes PfixQuickSortOfNumeric1DArray
'myZdouOrgData(i) : 元データ配列
    myXdouMax = Empty: myXdouMin = Empty: myXdouMedian = Empty
  Dim myXlonLbnd As Long: myXlonLbnd = LBound(myZdouOrgData)
  Dim myXlonUbnd As Long: myXlonUbnd = UBound(myZdouOrgData)
  Dim myXlonCnt As Long: myXlonCnt = myXlonUbnd - myXlonLbnd + 1
  Dim i As Long, myZdouTmp() As Double
    ReDim Preserve myZdouTmp(myXlonUbnd - myXlonLbnd + 1) As Double
    For i = myXlonLbnd To myXlonUbnd
        myZdouTmp(i) = myZdouOrgData(i)
    Next i
    Call PfixQuickSortOfNumeric1DArray(myZdouTmp, myXlonLbnd, myXlonUbnd, 1)
    myXdouMax = myZdouTmp(myXlonUbnd)
    myXdouMin = myZdouTmp(myXlonLbnd)
    If myXlonCnt Mod 2 = 1 Then
        myXdouMedian = myZdouTmp(Int((myXlonLbnd + myXlonUbnd) / 2))
    Else
      Dim m As Long, n As Long, myXdouM1 As Double, myXdouM2 As Double
        m = Int((myXlonLbnd + myXlonUbnd) / 2): n = m + 1
        myXdouM1 = myZdouTmp(m): myXdouM2 = myZdouTmp(n)
        myXdouMedian = (myXdouM1 + myXdouM2) / 2
    End If
    Erase myZdouTmp
End Sub

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

 '定型Ｆ_配列計算-標準偏差
Private Function PfncdouStDevP(ByRef myZdouOrgData() As Double) As Double
'myZdouOrgData(i) : 元データ配列
    PfncdouStDevP = Empty
  Dim myXdouSumSq As Double: myXdouSumSq = 0
  Dim myXdouSum As Double: myXdouSum = 0
  Dim i As Long, n As Long: n = 0
    For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
        n = n + 1
        myXdouSumSq = myXdouSumSq + (myZdouOrgData(i)) ^ 2
        myXdouSum = myXdouSum + myZdouOrgData(i)
    Next i
    If n <= 0 Then Exit Function
    PfncdouStDevP = ((myXdouSumSq - (myXdouSum ^ 2) / n) / n) ^ 0.5
End Function

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

 '定型Ｆ_1次元配列で数値を検索して最初のデータ位置を取得する
Private Function PfnclonMatchNumeric( _
            ByVal myXdouIfCndtn As Double, ByRef myZdouOrgData() As Double, _
            Optional ByVal coXbisNearest As Boolean = False) As Long
'myZdouOrgData(i): 元データ配列
'coXbisNearest = False : 指定数値と一致する値
'coXbisNearest = True  : 指定数値に最も近い値
    PfnclonMatchNumeric = Empty
  Dim myXdouTmp As Double, L As Long
    On Error GoTo ExitPath
    L = LBound(myZdouOrgData): myXdouTmp = myZdouOrgData(L)
    On Error GoTo 0
  Dim i As Long, n As Long: n = L - 1
    If coXbisNearest = True Then
      Dim myXdouPrvs As Double, myXdouPrsnt As Double, myXlonSgn As Long: myXlonSgn = 1
        myXdouTmp = myZdouOrgData(L) - myXdouIfCndtn
        If myXdouTmp < 0 Then myXlonSgn = -1
        For i = LBound(myZdouOrgData) + 1 To UBound(myZdouOrgData)
            myXdouTmp = myZdouOrgData(i) - myXdouIfCndtn
            If myXlonSgn * myXdouTmp < 0 Then
                n = i
                myXdouPrvs = myZdouOrgData(i - 1) - myXdouIfCndtn
                myXdouPrsnt = myZdouOrgData(i) - myXdouIfCndtn
                If Abs(myXdouPrsnt) >= Abs(myXdouPrvs) Then n = i - 1
                Exit For
            End If
        Next i
    Else
        For i = LBound(myZdouOrgData) To UBound(myZdouOrgData)
            myXdouTmp = myZdouOrgData(i) - myXdouIfCndtn
            If myXdouTmp = 0 Then
                n = i: Exit For
            End If
        Next i
    End If
    PfnclonMatchNumeric = n
ExitPath:
End Function

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_1次元の数値配列内のデータから統計量を計算する
'Private Sub instCBscSttstcs1Din()
'  Dim myXlonDataCnt As Long, myXdouSum As Double
'  Dim myXdouMax As Double, myXlonMaxDataNo As Long
'  Dim myXdouMin As Double, myXlonMinDataNo As Long
'  Dim myXdouMaxAbs As Double, myXlonMaxAbsDataNo As Long
'  Dim myXdouMedian As Double, myXdouArthmtcMean As Double
'  Dim myXdouStDev As Double
''//クラス内変数への入力
'  Dim myXlonBgn As Long, myXlonEnd As Long, myXlonCalStep As Long
'    myXlonBgn = 1
'    myXlonEnd = 0
'    myXlonCalStep = 1
'  Dim myXlonOrgCnt As Long
'    myXlonOrgCnt = 1
'  Dim myZdouOrgData() As Double
'    ReDim myZdouOrgData(myXlonOrgCnt) As Double
'    myZdouOrgData(myXlonOrgCnt) = 1
'  Dim myXstrOutOptn As String: myXstrOutOptn = "11111111111"
'    'myXstrOutOptn = "10000000000" : N
'    'myXstrOutOptn = "01000000000" : Sum
'    'myXstrOutOptn = "00100000000" : Max
'    'myXstrOutOptn = "00010000000" : MaxDataNo
'    'myXstrOutOptn = "00001000000" : Min
'    'myXstrOutOptn = "00000100000" : MinDataNo
'    'myXstrOutOptn = "00000010000" : MaxAbs
'    'myXstrOutOptn = "00000001000" : MaxAbsDataNo
'    'myXstrOutOptn = "00000000100" : Median
'    'myXstrOutOptn = "00000000010" : ArithmeticMean
'    'myXstrOutOptn = "00000000001" : StDev
''//CBscSttstcs1Dinクラスモジュールのインスタンス生成
'  Dim i As Long
'  Dim myXinsBscSttstcs1Din As CBscSttstcs1Din
'    Set myXinsBscSttstcs1Din = New CBscSttstcs1Din
'    With myXinsBscSttstcs1Din
'    '//クラス内変数への入力
'        .letOrgData = myZdouOrgData
'        .letBgn = myXlonBgn
'        .letEnd = myXlonEnd
'        .letCalStep = myXlonCalStep
'        .letOutOptn = myXstrOutOptn
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonDataCnt = .getDataCnt
'        myXdouSum = .getSum
'        myXdouMax = .getMax
'        myXlonMaxDataNo = .getMaxDataNo
'        myXdouMin = .getMin
'        myXlonMinDataNo = .getMinDataNo
'        myXdouMaxAbs = .getMaxAbs
'        myXlonMaxAbsDataNo = .getMaxAbsDataNo
'        myXdouMedian = .getMedian
'        myXdouArthmtcMean = .getArthmtcMean
'        myXdouStDev = .getStDev
'    End With
'JumpPath:
'    Set myXinsBscSttstcs1Din = Nothing
'End Sub
