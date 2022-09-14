VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSrchShtCmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes CSrchSht
'Includes CSrchCmnt
'Includes PfnclonArrayDimension

Option Explicit
Option Base 1

'◆ClassProc名_シート上とコメントのデータから文字列を検索してデータと位置情報を取得する
'Rev.005

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSrchShtCmnt"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonValCnt As Long, myZstrVal() As String, myZlonPstn() As Long, _
            myZobjPstnRng() As Object
  
'//PropertySetLet_入力データ
  Private myXobjSrchSheet As Object, myZvarSrchCndtn As Variant
  Private myXbisInStrOptn As Boolean
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private Lo As Long
  Private myZlonValSrchCnt As Long, myZvarValSrchCndtn() As Variant
  Private myZlonCmntSrchCnt As Long, myZvarCmntSrchCndtn() As Variant
  Private myXlonValCmntChck() As Long
  Private myXlonTmpValCnt As Long, myZvarTmpVal() As Variant, _
            myZlonTmpValPstn() As Long, myZobjTmpValPstnRng() As Object
  Private myXlonTmpCmntCnt As Long, myZvarTmpCmnt() As Variant, _
            myZlonTmpCmntPstn() As Long, myZobjTmpCmntPstnRng() As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Lo = Empty
    myZlonValSrchCnt = Empty: Erase myZvarValSrchCndtn
    myZlonCmntSrchCnt = Empty: Erase myZvarCmntSrchCndtn
    Erase myXlonValCmntChck
    myXlonTmpValCnt = Empty: Erase myZvarTmpVal
    Erase myZlonTmpValPstn: Erase myZobjTmpValPstnRng
    myXlonTmpCmntCnt = Empty: Erase myZvarTmpCmnt
    Erase myZlonTmpCmntPstn: Erase myZobjTmpCmntPstnRng
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getValCnt() As Long
    getValCnt = myXlonValCnt
End Property

Public Property Get getValAry(ByVal i As Long) As String
    getValAry = myZstrVal(i)
'myZstrVal(i) : 取得文字列
End Property

Public Property Get getPstnAry(ByVal i As Long, ByVal j As Long) As Long
    getPstnAry = myZlonPstn(i, j)
'myZlonPstn(i, j) : 行列位置
End Property

Public Property Get getPstnRngAry(ByVal i As Long) As Object
    Set getPstnRngAry = myZobjPstnRng(i)
'myZobjPstnRng(i) : 行列位置のセル
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonValCnt = Empty: Erase myZstrVal: Erase myZlonPstn: Erase myZobjPstnRng
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Set setSrchSheet(ByVal RHS As Object)
    Set myXobjSrchSheet = RHS
End Property

Public Property Let letSrchCndtn(ByVal RHS As Variant)
    myZvarSrchCndtn = RHS
'myZvarSrchCndtn(i, 1) : 検索文字列
'myZvarSrchCndtn(i, 2) : オフセット行数
'myZvarSrchCndtn(i, 3) : オフセット列数
'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
End Property

Public Property Let letInStrOptn(ByVal RHS As Boolean)
    myXbisInStrOptn = RHS
'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXobjSrchSheet Is Nothing Then GoTo ExitPath
    
'//検索条件の配列を設定
  Dim myZlonTmp(1) As Long: Lo = LBound(myZlonTmp)

'//配列変数の次元数を取得
  Dim myXlonAryIndx As Long
    myXlonAryIndx = PfnclonArrayDimension(myZvarSrchCndtn)
    
'//配列変数の次元数に応じて処理を実施
  Dim Li As Long, Ui1 As Long, Ui2 As Long, m As Long, n As Long
  Dim i As Long, j As Long, k As Long
    If myXlonAryIndx = 1 Then
        Li = LBound(myZvarSrchCndtn): Ui1 = UBound(myZvarSrchCndtn)
        m = 1: n = Ui1 - Li + 1
        If n <> 4 Then GoTo ExitPath
        
        If myZvarSrchCndtn(Li + 3) = 0 Then
            k = 1
            ReDim myZvarValSrchCndtn(Lo, Lo + 3) As Variant
            myZvarValSrchCndtn(Lo, Lo + 0) = myZvarSrchCndtn(Li + 0)
            myZvarValSrchCndtn(Lo, Lo + 1) = myZvarSrchCndtn(Li + 1)
            myZvarValSrchCndtn(Lo, Lo + 2) = myZvarSrchCndtn(Li + 2)
        Else
            ReDim myZvarCmntSrchCndtn(Lo, Lo + 3) As Variant
            myZvarCmntSrchCndtn(Lo, Lo + 0) = myZvarSrchCndtn(Li + 0)
            myZvarCmntSrchCndtn(Lo, Lo + 1) = myZvarSrchCndtn(Li + 1)
            myZvarCmntSrchCndtn(Lo, Lo + 2) = myZvarSrchCndtn(Li + 2)
        End If
        myZlonValSrchCnt = k
        myZlonCmntSrchCnt = m - k
        
    ElseIf myXlonAryIndx = 2 Then
        Li = LBound(myZvarSrchCndtn, 1): Ui1 = UBound(myZvarSrchCndtn, 1)
        Ui2 = UBound(myZvarSrchCndtn, 2)
        m = Ui1 - Li + 1: n = Ui2 - Li + 1
        If n <> 4 Then GoTo ExitPath
        
        i = m + Lo - 1: j = n + Lo - 1: k = 0
        ReDim myXlonValCmntChck(i) As Long
        For i = Li To Ui1
            myXlonValCmntChck(i + Lo - Li) = myZvarSrchCndtn(i, Li + 3)
            If myXlonValCmntChck(i + Lo - Li) = 0 Then k = k + 1
        Next i
        myZlonValSrchCnt = k
        myZlonCmntSrchCnt = m - k
        
        If myZlonValSrchCnt > 0 Then
            i = myZlonValSrchCnt + Lo - 1
            ReDim myZvarValSrchCndtn(i, j - 1) As Variant
        End If
        If myZlonCmntSrchCnt > 0 Then
            i = myZlonCmntSrchCnt + Lo - 1
            ReDim myZvarCmntSrchCndtn(i, j - 1) As Variant
        End If
        m = Lo - 1: n = Lo - 1
        For i = Li To Ui1
            If myZvarSrchCndtn(i, Li + 3) = 0 Then
                m = m + 1
                myZvarValSrchCndtn(m, Lo + 0) = myZvarSrchCndtn(i, Li + 0)
                myZvarValSrchCndtn(m, Lo + 1) = myZvarSrchCndtn(i, Li + 1)
                myZvarValSrchCndtn(m, Lo + 2) = myZvarSrchCndtn(i, Li + 2)
            Else
                n = n + 1
                myZvarCmntSrchCndtn(n, Lo + 0) = myZvarSrchCndtn(i, Li + 0)
                myZvarCmntSrchCndtn(n, Lo + 1) = myZvarSrchCndtn(i, Li + 1)
                myZvarCmntSrchCndtn(n, Lo + 2) = myZvarSrchCndtn(i, Li + 2)
            End If
        Next i
        
    Else
        GoTo ExitPath
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
        
'//シート上の文字列検索と位置の取得
    Call instCSrchSht
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//シート上のコメント内容検索と位置の取得
    Call instCSrchCmnt
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//出力配列変数に格納
    myXlonValCnt = myXlonTmpValCnt + myXlonTmpCmntCnt
    If myXlonValCnt <> myZlonValSrchCnt + myZlonCmntSrchCnt Then GoTo ExitPath
    
  Dim i As Long, j As Long, k As Long, m As Long, n As Long
    i = myXlonValCnt + Lo - 1: j = Lo + 1
    ReDim myZstrVal(i) As String
    ReDim myZlonPstn(i, j) As Long
    ReDim myZobjPstnRng(i) As Object
    k = Lo - 1: m = Lo - 1: n = Lo - 1
    On Error Resume Next
    For i = LBound(myXlonValCmntChck) To UBound(myXlonValCmntChck)
        If myXlonValCmntChck(i) = 0 Then
            k = k + 1: m = m + 1
            myZstrVal(k) = myZvarTmpVal(m)
            myZlonPstn(k, Lo + 0) = myZlonTmpValPstn(m, Lo + 0)
            myZlonPstn(k, Lo + 1) = myZlonTmpValPstn(m, Lo + 1)
            Set myZobjPstnRng(k) = myZobjTmpValPstnRng(m)
        Else
            k = k + 1: n = n + 1
            myZstrVal(k) = myZvarTmpCmnt(n)
            myZlonPstn(k, Lo + 0) = myZlonTmpCmntPstn(n, Lo + 0)
            myZlonPstn(k, Lo + 1) = myZlonTmpCmntPstn(n, Lo + 1)
            Set myZobjPstnRng(k) = myZobjTmpValPstnRng(n)
        End If
    Next i
    On Error GoTo 0
    
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

'◆ClassProc名_シート上のデータから文字列を検索してデータと位置情報を取得する
Private Sub instCSrchSht()
    myXbisExitFlag = False
    If myZlonValSrchCnt <= 0 Then Exit Sub
    
  Dim Lc As Long
  Dim i As Long, j As Long
  Dim myXinsSrchSht As CSrchSht: Set myXinsSrchSht = New CSrchSht
    With myXinsSrchSht
    '//文字列検索シートと検索条件を設定
        Set .setSrchSheet = myXobjSrchSheet
        .letSrchCndtn = myZvarValSrchCndtn
        .letInStrOptn = myXbisInStrOptn
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonTmpValCnt = .getValCnt
        If myXlonTmpValCnt <= 0 Then GoTo ExitPath
        If myXlonTmpValCnt < myZlonValSrchCnt Then GoTo ExitPath
        i = myXlonTmpValCnt + Lo - 1: j = Lo + 1
        ReDim myZvarTmpVal(i) As Variant
        ReDim myZlonTmpValPstn(i, j) As Long
        ReDim myZobjTmpValPstnRng(i) As Object
        Lc = .getOptnBase
        For i = 1 To myXlonTmpValCnt
            myZvarTmpVal(i + Lo - 1) = .getValAry(i + Lc - 1)
            myZlonTmpValPstn(i + Lo - 1, Lo + 0) = .getPstnAry(i + Lc - 1, Lc + 0)
            myZlonTmpValPstn(i + Lo - 1, Lo + 1) = .getPstnAry(i + Lc - 1, Lc + 1)
            Set myZobjTmpValPstnRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
        Next i
    End With
    Set myXinsSrchSht = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'◆ClassProc名_シート上のコメントから文字列を検索してデータと位置情報を取得する
Private Sub instCSrchCmnt()
    myXbisExitFlag = False
    If myZlonCmntSrchCnt <= 0 Then Exit Sub
    
  Dim Lc As Long
  Dim i As Long, j As Long
  Dim myXinsSrchCmnt As CSrchCmnt: Set myXinsSrchCmnt = New CSrchCmnt
    With myXinsSrchCmnt
    '//クラス内変数への入力
        Set .setSrchSheet = myXobjSrchSheet
        .letSrchCndtn = myZvarCmntSrchCndtn
        .letInStrOptn = myXbisInStrOptn
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonTmpCmntCnt = .getCmntCnt
        If myXlonTmpCmntCnt <= 0 Then GoTo ExitPath
        If myXlonTmpCmntCnt < myZlonCmntSrchCnt Then GoTo ExitPath
        i = myXlonTmpCmntCnt + Lo - 1: j = Lo + 1
        ReDim myZvarTmpCmnt(i) As Variant
        ReDim myZlonTmpCmntPstn(i, j) As Long
        ReDim myZobjTmpCmntPstnRng(i) As Object
        Lc = .getOptnBase
        For i = 1 To myXlonTmpCmntCnt
            myZvarTmpCmnt(i + Lo - 1) = .getCmntAry(i + Lc - 1)
            myZlonTmpCmntPstn(i + Lo - 1, Lo + 0) = .getPstnAry(i + Lc - 1, Lc + 0)
            myZlonTmpCmntPstn(i + Lo - 1, Lo + 1) = .getPstnAry(i + Lc - 1, Lc + 1)
            Set myZobjTmpCmntPstnRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
        Next i
    End With
    Set myXinsSrchCmnt = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

 '定型Ｆ_配列変数の次元数を取得する
Private Function PfnclonArrayDimension(ByRef myZvarDataAry As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarDataAry) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarDataAry, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
End Function

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上のデータから文字列を検索してデータと位置情報を取得する
'Private Sub instCSrchShtCmnt()
'  Dim Lc As Long
'  Dim myXlonValCnt As Long, myZstrVal() As String, myZlonPstn() As Long, _
'        myZobjPstnRng() As Object
'    'myZstrVal(i) : 取得文字列
'    'myZlonPstn(i, j) : 行列位置
'    'myZobjPstnRng(i) : 行列位置のセル
''//クラス内変数への入力
'  Dim myZlonBuf(1) As Long, Li As Long: Li = LBound(myZlonBuf)
'  Dim myXobjSrchSheet As Object: Set myXobjSrchSheet = ActiveSheet
'  Dim myZvarSrchCndtn As Variant: ReDim myZvarSrchCndtn(Li + 1, Li + 3) As Variant
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
'    myZvarSrchCndtn(Li + 0, Li + 0) = "エクセルシート上のデータを検索"
'    myZvarSrchCndtn(Li + 0, Li + 1) = 0
'    myZvarSrchCndtn(Li + 0, Li + 2) = 0
'    myZvarSrchCndtn(Li + 0, Li + 3) = 0
'    myZvarSrchCndtn(Li + 1, Li + 0) = "データ行"
'    myZvarSrchCndtn(Li + 1, Li + 1) = 0
'    myZvarSrchCndtn(Li + 1, Li + 2) = 0
'    myZvarSrchCndtn(Li + 1, Li + 3) = 1
'  Dim myXbisInStrOptn As Boolean: myXbisInStrOptn = False
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
''//CSrchShtCmntクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsSrchShtCmnt As CSrchShtCmnt: Set myXinsSrchShtCmnt = New CSrchShtCmnt
'    With myXinsSrchShtCmnt
'    '//文字列検索シートと検索条件を設定
'        Set .setSrchSheet = myXobjSrchSheet
'        .letSrchCndtn = myZvarSrchCndtn
'        .letInStrOptn = myXbisInStrOptn
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonValCnt = .getValCnt
'        If myXlonValCnt <= 0 Then GoTo JumpPath
'        i = myXlonValCnt + Lo - 1: j = Lo + 1
'        ReDim myZstrVal(i) As String
'        ReDim myZlonPstn(i, j) As Long
'        ReDim myZobjPstnRng(i) As Object
'        Lc = .getOptnBase
'        For i = 1 To myXlonValCnt
'            myZstrVal(i + Lo - 1) = .getValAry(i + Lc - 1)
'            myZlonPstn(i + Lo - 1, Lo + 0) = .getPstnAry(i + Lc - 1, Lc + 0)
'            myZlonPstn(i + Lo - 1, Lo + 1) = .getPstnAry(i + Lc - 1, Lc + 1)
'            Set myZobjPstnRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
'        Next i
'    End With
'JumpPath:
'    Set myXinsSrchShtCmnt = Nothing
'    Set myXobjSrchSheet = Nothing: myZvarSrchCndtn = Empty
'    Call variablesOfCSrchShtCmnt(myXlonValCnt, myZstrVal)    'Debug.Print
'End Sub
'Private Sub variablesOfCSrchShtCmnt( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSrchShtCmntクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
