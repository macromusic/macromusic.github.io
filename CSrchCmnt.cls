VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSrchCmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfnclonArrayDimension
'Includes PincSearchCommentCellMulti
'Includes PfncobjSearchCommentCell

Option Explicit
Option Base 1

'◆ClassProc名_シート上のコメントから文字列を検索してデータと位置情報を取得する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSrchCmnt"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonCmntCnt As Long, myZstrCmnt() As String, myZlonPstn() As Long, _
            myZobjPstnRng() As Object
  
'//PropertySetLet_入力データ
  Private myXobjSrchSheet As Object, myZvarSrchCndtn As Variant
  Private myXbisInStrOptn As Boolean
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private Lo As Long
  Private myZvarSubSrchCndtn() As Variant, myZvarCmntPstn() As Variant

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Lo = Empty
    Erase myZvarSubSrchCndtn: Erase myZvarCmntPstn
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getCmntCnt() As Long
    getCmntCnt = myXlonCmntCnt
End Property

Public Property Get getCmntAry(ByVal i As Long) As String
    getCmntAry = myZstrCmnt(i)
'myZstrCmnt(i) : 取得文字列
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
    myXlonCmntCnt = Empty: Erase myZstrCmnt: Erase myZlonPstn: Erase myZobjPstnRng
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
  Dim Li As Long, Ui1 As Long, Ui2 As Long, n As Long, i As Long, j As Long
    If myXlonAryIndx = 1 Then
        Li = LBound(myZvarSrchCndtn): Ui1 = UBound(myZvarSrchCndtn)
        n = Ui1 - Li + 1
        If n <> 3 Then GoTo ExitPath
        ReDim myZvarSubSrchCndtn(Lo, Lo + 2) As Variant
        myZvarSubSrchCndtn(Lo, Lo + 0) = myZvarSrchCndtn(Li + 0)
        myZvarSubSrchCndtn(Lo, Lo + 1) = myZvarSrchCndtn(Li + 1)
        myZvarSubSrchCndtn(Lo, Lo + 2) = myZvarSrchCndtn(Li + 2)
        
    ElseIf myXlonAryIndx = 2 Then
        Li = LBound(myZvarSrchCndtn, 1): Ui1 = UBound(myZvarSrchCndtn, 1)
        Ui2 = UBound(myZvarSrchCndtn, 2)
        n = Ui2 - Li + 1
        If n <> 3 Then GoTo ExitPath
        i = Ui1 + Lo - Li: j = Ui2 + Lo - Li
        ReDim myZvarSubSrchCndtn(i, j) As Variant
        For i = Li To Ui1
            myZvarSubSrchCndtn(i + Lo - Li, Lo + 0) = myZvarSrchCndtn(i, Li + 0)
            myZvarSubSrchCndtn(i + Lo - Li, Lo + 1) = myZvarSrchCndtn(i, Li + 1)
            myZvarSubSrchCndtn(i + Lo - Li, Lo + 2) = myZvarSrchCndtn(i, Li + 2)
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
    
'//S:シート上のコメントで複数の文字列を検索してそれぞれの最初のコメントと位置の一覧を取得
    Call PincSearchCommentCellMulti( _
            myXlonCmntCnt, myZvarCmntPstn, _
            myZvarSubSrchCndtn, myXobjSrchSheet, myXbisInStrOptn)
    If myXlonCmntCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//出力配列変数に格納
  Dim i As Long, j As Long, k As Long
    i = myXlonCmntCnt + Lo - 1: j = Lo + 1
    ReDim myZstrCmnt(i) As String
    ReDim myZlonPstn(i, j) As Long
    ReDim myZobjPstnRng(i) As Object
    For i = LBound(myZvarCmntPstn) To UBound(myZvarCmntPstn)
        myZstrCmnt(i) = myZvarCmntPstn(i, Lo + 0)
        myZlonPstn(i, Lo + 0) = myZvarCmntPstn(i, Lo + 1)
        myZlonPstn(i, Lo + 1) = myZvarCmntPstn(i, Lo + 2)
    Next i
  
    On Error Resume Next
    For k = LBound(myZlonPstn, 1) To UBound(myZlonPstn, 1)
        i = Empty: j = Empty
        i = myZlonPstn(k, Lo + 0)
        j = myZlonPstn(k, Lo + 1)
        Set myZobjPstnRng(k) = myXobjSrchSheet.Cells(i, j)
    Next k
    On Error GoTo 0
    
ExitPath:
    Call initializeModuleVariables
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

 '定型Ｐ_シート上のコメントで複数の文字列を検索してそれぞれの最初のコメントと位置の一覧を取得する
Private Sub PincSearchCommentCellMulti( _
            myXlonCmntCnt As Long, myZvarCmntPstn() As Variant, _
            ByRef myZvarSrchCndtn() As Variant, ByVal myXobjSrchSheet As Object, _
            Optional ByVal coXbisInStrOptn As Boolean = False)
'Includes PfncobjSearchCommentCell
'myZvarCmntPstn(i, 1) : 取得データ
'myZvarCmntPstn(i, 2) : 行位置
'myZvarCmntPstn(i, 3) : 列位置
'myZvarSrchCndtn(i, 1) : 検索文字列
'myZvarSrchCndtn(i, 2) : オフセット行数
'myZvarSrchCndtn(i, 3) : オフセット列数
'coXbisInStrOptn = False : 指定文字列と一致する条件
'coXbisInStrOptn = True  : 指定文字列を含む条件
    myXlonCmntCnt = Empty: Erase myZvarCmntPstn
  Dim myXstrTmp As String, Li As Long, Ui As Long
    On Error GoTo ExitPath
    Li = LBound(myZvarSrchCndtn, 1): myXstrTmp = myZvarSrchCndtn(Li, Li)
    Ui = UBound(myZvarSrchCndtn, 1)
    On Error GoTo 0
    myXlonCmntCnt = Ui - Li + 1
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long: i = myXlonCmntCnt + Lo - 1: ReDim myZvarCmntPstn(i, Lo + 2) As Variant
  Dim myXstrIfCndtn As String, myXlonROfst As Long, myXlonCOfst As Long
  Dim myXobjTmpRng As Object
    For i = Li To Ui
   '//検索条件を設定
        myXstrIfCndtn = CStr(myZvarSrchCndtn(i, Li + 0))
        myXlonROfst = CLng(myZvarSrchCndtn(i, Li + 1))
        myXlonCOfst = CLng(myZvarSrchCndtn(i, Li + 2))
        If myXstrIfCndtn = "" Then GoTo ExitPath
   '//シート上のコメント内容を検索してセルを取得
        Set myXobjTmpRng = PfncobjSearchCommentCell( _
                            myXstrIfCndtn, myXobjSrchSheet, coXbisInStrOptn)
   '//検索結果を出力変数に格納
        If Not myXobjTmpRng Is Nothing Then
            Set myXobjTmpRng = myXobjTmpRng.Offset(myXlonROfst, myXlonCOfst)
            myZvarCmntPstn(i, Lo + 0) = myXobjTmpRng.Comment.Text
            myZvarCmntPstn(i, Lo + 1) = myXobjTmpRng.Row
            myZvarCmntPstn(i, Lo + 2) = myXobjTmpRng.Column
        Else
            myZvarCmntPstn(i, Lo + 0) = ""
            myZvarCmntPstn(i, Lo + 1) = 0
            myZvarCmntPstn(i, Lo + 2) = 0
        End If
    Next i
    Set myXobjTmpRng = Nothing
ExitPath:
End Sub

 '定型Ｆ_シート上のコメント内容を検索してセルを取得する
Private Function PfncobjSearchCommentCell( _
            ByVal myXstrSrchVal As Variant, ByVal myXobjSrchSheet As Object, _
            Optional ByVal coXbisInStrOptn As Boolean = False) As Object
'coXbisInStrOptn = False : 指定文字列と一致する条件
'coXbisInStrOptn = True  : 指定文字列を含む条件
    Set PfncobjSearchCommentCell = Nothing
    If myXstrSrchVal = "" Then Exit Function
    If myXobjSrchSheet Is Nothing Then Exit Function
  Dim myXobjCmntCells As Object
    On Error GoTo ExitPath
    Set myXobjCmntCells = myXobjSrchSheet.Cells.SpecialCells(xlCellTypeComments)
    On Error GoTo 0
  Dim myXobjCmntRng As Object, myXstrTmp As String
    If coXbisInStrOptn = True Then
        For Each myXobjCmntRng In myXobjCmntCells
            myXstrTmp = myXobjCmntRng.Comment.Text
            If InStr(myXstrTmp, myXstrSrchVal) > 0 Then Exit For
        Next myXobjCmntRng
    Else
        For Each myXobjCmntRng In myXobjCmntCells
            myXstrTmp = myXobjCmntRng.Comment.Text
            If myXstrTmp = myXstrSrchVal Then Exit For
        Next myXobjCmntRng
    End If
    Set PfncobjSearchCommentCell = myXobjCmntRng
    Set myXobjCmntCells = Nothing: Set myXobjCmntRng = Nothing
ExitPath:
End Function

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上のコメントから文字列を検索してデータと位置情報を取得する
'Private Sub instCSrchCmnt()
'  Dim Lc As Long
'  Dim myXlonCmntCnt As Long, myZstrCmnt() As String, myZlonPstn() As Long, _
'        myZobjPstnRng() As Object
'    'myZstrCmnt(i) : 取得文字列
'    'myZlonPstn(i, j) : 行列位置
'    'myZobjPstnRng(i) : 行列位置のセル
''//クラス内変数への入力
'  Dim myZlonBuf(1) As Long, Li As Long: Li = LBound(myZlonBuf)
'  Dim myXobjSrchSheet As Object: Set myXobjSrchSheet = ActiveSheet
'  Dim myZvarSrchCndtn As Variant: ReDim myZvarSrchCndtn(Li, Li + 2) As Variant
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    myZvarSrchCndtn(Li, Li + 0) = "データ列"
'    myZvarSrchCndtn(Li, Li + 1) = 0
'    myZvarSrchCndtn(Li, Li + 2) = 0
'  Dim myXbisInStrOptn As Boolean: myXbisInStrOptn = False
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
''//CSrchCmntクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsSrchCmnt As CSrchCmnt: Set myXinsSrchCmnt = New CSrchCmnt
'    With myXinsSrchCmnt
'    '//クラス内変数への入力
'        Set .setSrchSheet = myXobjSrchSheet
'        .letSrchCndtn = myZvarSrchCndtn
'        .letInStrOptn = myXbisInStrOptn
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonCmntCnt = .getCmntCnt
'        If myXlonCmntCnt <= 0 Then GoTo JumpPath
'        i = myXlonCmntCnt + Lo - 1: j = Lo + 1
'        ReDim myZstrCmnt(i) As String
'        ReDim myZlonPstn(i, j) As Long
'        ReDim myZobjPstnRng(i) As Object
'        Lc = .getOptnBase
'        For i = 1 To myXlonCmntCnt
'            myZstrCmnt(i + Lo - 1) = .getCmntAry(i + Lc - 1)
'            myZlonPstn(i + Lo - 1, Lo + 0) = .getPstnAry(i + Lc - 1, Lc + 0)
'            myZlonPstn(i + Lo - 1, Lo + 1) = .getPstnAry(i + Lc - 1, Lc + 1)
'            Set myZobjPstnRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
'        Next i
'    End With
'JumpPath:
'    Set myXinsSrchCmnt = Nothing
'    Set myXobjSrchSheet = Nothing: myZvarSrchCndtn = Empty
'    Call variablesOfCSrchCmnt(myXlonCmntCnt, myZstrCmnt)    'Debug.Print
'ExitPath:
'End Sub
'Private Sub variablesOfCSrchCmnt( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSrchCmntクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
