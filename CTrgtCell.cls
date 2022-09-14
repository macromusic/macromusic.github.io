VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTrgtCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfnclonArrayDimension
'Includes PfncobjSearchValueCell
'Includes PfncobjSearchCommentCell

Option Explicit
Option Base 1

'◆ClassProc名_シート上のデータから文字列を検索してセルを取得する
'Rev.009
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CTrgtCell"
  
'//PropertyGet_出力データ
  Private myXobjTrgtCell As Object
  
'//PropertySetLet_入力制御信号
  Private myXbisCmntSrch As Boolean
  
'//PropertySetLet_入力データ
  Private myXobjSrchSheet As Object, myZvarSrchCndtn As Variant
  Private myXbisInStrOptn As Boolean
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXstrSrchVal As String, myXlonRowOfst As Long, myXlonColOfst As Long

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrSrchVal = Empty: myXlonRowOfst = Empty: myXlonColOfst = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getTrgtCell() As Object
    Set getTrgtCell = myXobjTrgtCell
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    Set myXobjTrgtCell = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letCmntSrch(ByVal RHS As Boolean)
    myXbisCmntSrch = RHS
'myXbisCmntSrch = True  : シート上のコメント内容を検索する
'myXbisCmntSrch = False : シート上の文字列を検索する
End Property

Public Property Set setSrchSheet(ByVal RHS As Object)
    Set myXobjSrchSheet = RHS
End Property

Public Property Let letSrchCndtn(ByVal RHS As Variant)
    myZvarSrchCndtn = RHS
'myZvarSrchCndtn(1) : 検索文字列
'myZvarSrchCndtn(2) : オフセット行数
'myZvarSrchCndtn(3) : オフセット列数
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
    
'//配列変数の次元数を取得
  Dim myXlonAryIndx As Long
    myXlonAryIndx = PfnclonArrayDimension(myZvarSrchCndtn)
    
'//配列変数の次元数に応じて処理を実施
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim Li As Long, Ui1 As Long, Ui2 As Long, n As Long, i As Long, j As Long
    If myXlonAryIndx = 1 Then
        Li = LBound(myZvarSrchCndtn): Ui1 = UBound(myZvarSrchCndtn)
        n = Ui1 - Li + 1
        If n <> 3 Then GoTo ExitPath
        myXstrSrchVal = myZvarSrchCndtn(Li + 0)
        myXlonRowOfst = myZvarSrchCndtn(Li + 1)
        myXlonColOfst = myZvarSrchCndtn(Li + 2)
        
    ElseIf myXlonAryIndx = 2 Then
        Li = LBound(myZvarSrchCndtn, 1): Ui1 = UBound(myZvarSrchCndtn, 1)
        Ui2 = UBound(myZvarSrchCndtn, 2)
        n = Ui2 - Li + 1
        If n <> 3 Then GoTo ExitPath
        myXstrSrchVal = myZvarSrchCndtn(Li, Li + 0)
        myXlonRowOfst = myZvarSrchCndtn(Li, Li + 1)
        myXlonColOfst = myZvarSrchCndtn(Li, Li + 2)
        
    Else
        GoTo ExitPath
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

'PublicF_
Public Function fncobjTargetCell() As Object
    Set fncobjTargetCell = Nothing
    Call ctrProc
    Set fncobjTargetCell = myXobjTrgtCell
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
        
'//C:文字列の検索方法を選択
    If myXbisCmntSrch = False Then
    '//S:シート上の文字列を検索してセルを取得
        Set myXobjTrgtCell = PfncobjSearchValueCell( _
                                myXstrSrchVal, myXobjSrchSheet, myXbisInStrOptn)
        If myXobjTrgtCell Is Nothing Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "2"    'PassFlag
        
    Else
    '//S:シート上のコメント内容を検索してセルを取得
        Set myXobjTrgtCell = PfncobjSearchCommentCell( _
                                myXstrSrchVal, myXobjSrchSheet, myXbisInStrOptn)
        If myXobjTrgtCell Is Nothing Then GoTo ExitPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "3"    'PassFlag
        
    End If
    
    On Error Resume Next
    Set myXobjTrgtCell = myXobjTrgtCell.Offset(myXlonRowOfst, myXlonColOfst)
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

 '定型Ｆ_シート上の文字列を検索してセルを取得する
Private Function PfncobjSearchValueCell( _
            ByVal myXstrSrchVal As String, ByVal myXobjSheet As Object, _
            Optional ByVal coXbisInStrOptn As Boolean = False) As Object
'coXbisInStrOptn = False : 指定文字列と一致する条件
'coXbisInStrOptn = True  : 指定文字列を含む条件
    Set PfncobjSearchValueCell = Nothing
    If myXstrSrchVal = "" Then Exit Function
    If myXobjSheet Is Nothing Then Exit Function
  Dim myXobjDataRng As Object
    Set myXobjDataRng = myXobjSheet.Range( _
            myXobjSheet.Cells(1, 1), _
            myXobjSheet.Cells.SpecialCells(xlCellTypeLastCell))
  Dim i As Long, j As Long
    i = myXobjDataRng.Rows.Count: j = myXobjDataRng.Columns.Count
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZvarRngData As Variant
    If i * j = 1 Then
        ReDim myZvarRngData(L, L) As Variant
        myZvarRngData(L, L) = myXobjDataRng.Value
    Else
        myZvarRngData = myXobjDataRng.Value
    End If
  Dim myXobjTrgtRng As Object, myXstrTmp As String
    On Error Resume Next
    If coXbisInStrOptn = True Then
        For j = LBound(myZvarRngData, 2) To UBound(myZvarRngData, 2)
            For i = LBound(myZvarRngData, 1) To UBound(myZvarRngData, 1)
                myXstrTmp = myZvarRngData(i, j)
                If InStr(myXstrTmp, myXstrSrchVal) > 0 Then
                    Set myXobjTrgtRng = myXobjSheet.Cells(i, j)
                    GoTo JumpPath
                End If
            Next i
        Next j
    Else
        For j = LBound(myZvarRngData, 2) To UBound(myZvarRngData, 2)
            For i = LBound(myZvarRngData, 1) To UBound(myZvarRngData, 1)
                myXstrTmp = myZvarRngData(i, j)
                If myXstrTmp = myXstrSrchVal Then
                    Set myXobjTrgtRng = myXobjSheet.Cells(i, j)
                    GoTo JumpPath
                End If
            Next i
        Next j
    End If
    On Error GoTo 0
JumpPath:
    Set PfncobjSearchValueCell = myXobjTrgtRng
    Set myXobjDataRng = Nothing: Set myXobjTrgtRng = Nothing
    myZvarRngData = Empty
End Function

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

''◆ClassProc名_シート上のデータから文字列を検索してセルを取得する
'Private Sub instCTrgtCell()
'  Dim myXobjTrgtCell As Object
''//クラス内変数への入力
'  Dim myXbisCmntSrch As Boolean: myXbisCmntSrch = False
'    'myXbisCmntSrch = False : シート上の文字列を検索する
'    'myXbisCmntSrch = True  : シート上のコメント内容を検索する
'  Dim myZlonBuf(1) As Long, Li As Long: Li = LBound(myZlonBuf)
'  Dim myXobjSrchSheet As Object: Set myXobjSrchSheet = ActiveSheet
'  Dim myZvarSrchCndtn As Variant: ReDim myZvarSrchCndtn(Li, Li + 2) As Variant
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    myZvarSrchCndtn(Li, Li + 0) = "エクセルシート上のデータを検索"
'    myZvarSrchCndtn(Li, Li + 1) = 0
'    myZvarSrchCndtn(Li, Li + 2) = 0
'  Dim myXbisInStrOptn As Boolean: myXbisInStrOptn = False
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
''//CTrgtCellクラスモジュールのインスタンス生成
'  Dim myXinsTrgtCell As CTrgtCell: Set myXinsTrgtCell = New CTrgtCell
'    With myXinsTrgtCell
'    '//クラス内変数への入力
'        .letCmntSrch = myXbisCmntSrch
'        Set .setSrchSheet = myXobjSrchSheet
'        .letSrchCndtn = myZvarSrchCndtn
'        .letInStrOptn = myXbisInStrOptn
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        Set myXobjTrgtCell = .fncobjTargetCell
'    End With
'    Set myXinsTrgtCell = Nothing
'    Set myXobjSrchSheet = Nothing: myZvarSrchCndtn = Empty
'End Sub
