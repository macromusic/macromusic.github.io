VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSeriesData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfixGetSheetRangeData
'Includes PfixSeriesData

Option Explicit
Option Base 1

'◆ClassProc名_シート上の連続するデータ範囲を取得する
'Rev.005

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSeriesData"

'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisRowDrctn As Boolean

'//PropertySetLet_入力データ
  Private myXobjSrchSheet As Object
  Private myXlonBgnRow As Long, myXlonBgnCol As Long

'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean

'//モジュール内変数_データ
  Private myXobjFrstCell As Object, myXobjLastCell As Object
  Private myXlonRowCnt As Long, myXlonColCnt As Long, myZstrShtData() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZstrShtData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getSrsDataCnt() As Long
    getSrsDataCnt = myXlonSrsDataCnt
End Property

Public Property Get getSrsDataAry(ByVal k As Long) As String
    getSrsDataAry = myZstrSrsData(k)
'myZstrSrsData(k) : 取得文字列
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Set setSrchSheet(ByVal RHS As Object)
    Set myXobjSrchSheet = RHS
End Property

Public Property Let letBgnRowCol(ByVal k As Long, ByVal RHS As Long)
    Select Case k
        Case 1: myXlonBgnRow = RHS
        Case 2: myXlonBgnCol = RHS
        Case Else
    End Select
End Property

Public Property Let letRowDrctn(ByVal RHS As Boolean)
    myXbisRowDrctn = RHS
'myXbisRowDrctn = True  : 行方向のみを検索
'myXbisRowDrctn = False : 列方向のみを検索
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False

    If myXobjSrchSheet Is Nothing Then GoTo ExitPath
    If myXlonBgnRow <= 0 Or myXlonBgnCol <= 0 Then GoTo ExitPath

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

'//シート上の指定範囲またはデータ全範囲のデータを配列に取込む
    Call PfixGetSheetRangeData(myXlonRowCnt, myXlonColCnt, myZstrShtData, _
                                myXobjSrchSheet, myXobjFrstCell, myXobjLastCell)
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
  
'//2次元配列内の任意の始点から行または列方向に連続する範囲のデータを取得
    Call PfixSeriesData(myXlonSrsDataCnt, myZstrSrsData, _
                        myZstrShtData, myXlonBgnRow, myXlonBgnCol, myXbisRowDrctn)
    If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

 '定型Ｐ_シート上の指定範囲またはデータ全範囲のデータを配列に取込む
Private Sub PfixGetSheetRangeData( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZstrShtData() As String, _
            ByVal myXobjSheet As Object, _
            ByVal myXobjFrstCell As Object, ByVal myXobjLastCell As Object)
'myZstrShtData(i, j) : データ
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZstrShtData
    If myXobjSheet Is Nothing Then Exit Sub
'//シート上の指定範囲をオブジェクト配列に取込む
  Dim myXobjShtRng As Object
    If myXobjFrstCell Is Nothing Then Set myXobjFrstCell = myXobjSheet.Cells(1, 1)
    If myXobjLastCell Is Nothing Then _
        Set myXobjLastCell = myXobjSheet.Cells.SpecialCells(xlCellTypeLastCell)
    Set myXobjShtRng = myXobjSheet.Range(myXobjFrstCell, myXobjLastCell)
    myXlonRowCnt = myXobjShtRng.Rows.Count
    myXlonColCnt = myXobjShtRng.Columns.Count
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then Exit Sub
'//オブジェクト配列からデータを取得
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZvarRngData As Variant, i As Long, j As Long
    i = myXlonRowCnt + L - 1: j = myXlonColCnt + L - 1
    ReDim myZstrShtData(i, j) As String
    myZvarRngData = myXobjShtRng.Value
    On Error Resume Next
    If myXlonRowCnt * myXlonColCnt = 1 Then
        myZstrShtData(L, L) = myZvarRngData
    Else
        For j = LBound(myZvarRngData, 2) To UBound(myZvarRngData, 2)
            For i = LBound(myZvarRngData, 1) To UBound(myZvarRngData, 1)
                myZstrShtData(i, j) = myZvarRngData(i, j)
            Next i
        Next j
    End If
    On Error GoTo 0
    Set myXobjShtRng = Nothing: myZvarRngData = Empty
End Sub

 '定型Ｐ_2次元配列内の任意の始点から行または列方向に連続する範囲のデータを取得する
Private Sub PfixSeriesData( _
            myXlonSrsDataCnt As Long, myZstrSrsData() As String, _
            ByRef myZstrOrgData() As String, _
            ByVal myXlonBgnRow As Long, ByVal myXlonBgnCol As Long, _
            Optional ByVal coXbisRowDrctn As Boolean = True)
'myZstrSrsData(i) : 連続範囲の取得データ
'myZdouOrgData(i, j) : 元データ配列
'coXbisRowDrctn = True  : 行方向のみを検索
'coXbisRowDrctn = False : 列方向のみを検索
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    If myXlonBgnRow < LBound(myZstrOrgData, 1) _
        Or myXlonBgnRow > UBound(myZstrOrgData, 1) Then Exit Sub
    If myXlonBgnCol < LBound(myZstrOrgData, 2) _
        Or myXlonBgnCol > UBound(myZstrOrgData, 2) Then Exit Sub
    myXstrTmp = myZstrOrgData(myXlonBgnRow, myXlonBgnCol)
    Li = LBound(myZstrOrgData, 1)
    On Error GoTo 0
  Dim myXlonEndRow As Long, myXlonEndCol As Long, i As Long, j As Long
  Dim n As Long: n = Li - 1
    If coXbisRowDrctn = True Then
    '//行方向
        If myXstrTmp = "" Then Exit Sub
        For i = myXlonBgnRow To UBound(myZstrOrgData, 1)
            If myZstrOrgData(i, myXlonBgnCol) = "" Then Exit For
            n = n + 1: ReDim Preserve myZstrSrsData(n) As String
            myZstrSrsData(n) = myZstrOrgData(i, myXlonBgnCol)
        Next i
    Else
    '//列方向
        If myXstrTmp = "" Then Exit Sub
        For j = myXlonBgnCol To UBound(myZstrOrgData, 2)
            If myZstrOrgData(myXlonBgnRow, j) = "" Then Exit For
            n = n + 1: ReDim Preserve myZstrSrsData(n) As String
            myZstrSrsData(n) = myZstrOrgData(myXlonBgnRow, j)
        Next j
    End If
    myXlonSrsDataCnt = n - Li + 1
ExitPath:
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上の連続するデータ範囲を取得する
'Private Sub instCSeriesData()
'  Dim Lc As Long
'  Dim myXlonSrsDataCnt As Long, myZstrSrsData() As String
'    'myZstrSrsData(k) : 取得文字列
''//クラス内変数への入力
'  Dim myXbisRowDrctn As Boolean: myXbisRowDrctn = True
'    'myXbisRowDrctn = True  : 行方向のみを検索
'    'myXbisRowDrctn = False : 列方向のみを検索
'  Dim myXlonBgnRow As Long: myXlonBgnRow = 8
'  Dim myXlonBgnCol As Long: myXlonBgnCol = 2
'  Dim myXobjSrchSheet As Object: Set myXobjSrchSheet = ActiveSheet
''//CSeriesDataクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim k As Long
'  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
'    With myXinsSeriesData
'    '//クラス内変数への入力
'        Set .setSrchSheet = myXobjSrchSheet
'        .letBgnRowCol(1) = myXlonBgnRow
'        .letBgnRowCol(2) = myXlonBgnCol
'        .letRowDrctn = True
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonSrsDataCnt = .getSrsDataCnt
'        If myXlonSrsDataCnt <= 0 Then GoTo JumpPath
'        k = myXlonSrsDataCnt + Lo - 1
'        ReDim myZstrSrsData(k) As String
'        Lc = .getOptnBase
'        For k = 1 To myXlonSrsDataCnt
'            myZstrSrsData(k + Lo - 1) = .getSrsDataAry(k + Lc - 1)
'        Next k
'    End With
'JumpPath:
'    Set myXinsSeriesData = Nothing
'    Set myXobjSrchSheet = Nothing
'    Call variablesOfCSeriesData(myXlonSrsDataCnt, myZstrSrsData) 'Debug.Print
'End Sub
'Private Sub variablesOfCSeriesData( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSeriesDataクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
