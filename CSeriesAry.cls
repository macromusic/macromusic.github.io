VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSeriesAry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfixGetSheetRangeData
'Includes PfixSeriesDataArray

Option Explicit
Option Base 1

'◆ClassProc名_シート上の連続するデータ範囲を行列で取得する
'Rev.005

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSeriesAry"

'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsData() As String

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

Public Property Get getSrsRowCnt() As Long
    getSrsRowCnt = myXlonSrsRowCnt
End Property

Public Property Get getSrsColCnt() As Long
    getSrsColCnt = myXlonSrsColCnt
End Property

Public Property Get getSrsDataAry(ByVal i As Long, ByVal j As Long) As String
    getSrsDataAry = myZstrSrsData(i, j)
'myZstrSrsData(i, j) : 取得文字列
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonSrsRowCnt = Empty: myXlonSrsColCnt = Empty: Erase myZstrSrsData
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
  
'//2次元配列内の任意の始点から行と列方向に連続する範囲のデータを取得
    Call PfixSeriesDataArray(myXlonSrsRowCnt, myXlonSrsColCnt, myZstrSrsData, _
                                myZstrShtData, myXlonBgnRow, myXlonBgnCol)
    If myXlonSrsRowCnt <= 0 Or myXlonSrsColCnt <= 0 Then GoTo ExitPath
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

 '定型Ｐ_2次元配列内の任意の始点から行と列方向に連続する範囲のデータを取得する
Private Sub PfixSeriesDataArray( _
            myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsData() As String, _
            ByRef myZstrOrgData() As String, _
            ByVal myXlonBgnRow As Long, ByVal myXlonBgnCol As Long)
'myZstrSrsData(i, j) : 連続範囲の取得データ
'myZdouOrgData(i, j) : 元データ配列
    myXlonSrsRowCnt = Empty: myXlonSrsColCnt = Empty: Erase myZstrSrsData
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
    If myXstrTmp = "" Then Exit Sub
'//行方向
    For i = myXlonBgnRow To UBound(myZstrOrgData, 1)
        If myZstrOrgData(i, myXlonBgnCol) = "" Then Exit For
    Next i
    myXlonEndRow = i - 1
'//列方向
    For j = myXlonBgnCol To UBound(myZstrOrgData, 2)
        If myZstrOrgData(myXlonBgnRow, j) = "" Then Exit For
    Next j
    myXlonEndCol = j - 1
'//連続範囲のデータを配列変数に格納
    If myXlonEndRow < myXlonBgnRow Or myXlonEndCol < myXlonBgnCol Then Exit Sub
    myXlonSrsRowCnt = myXlonEndRow - myXlonBgnRow + 1
    myXlonSrsColCnt = myXlonEndCol - myXlonBgnCol + 1
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
    i = myXlonSrsRowCnt + Lo - 1: j = myXlonSrsColCnt + Lo - 1
    ReDim myZstrSrsData(i, j) As String
  Dim m As Long, n As Long: n = Lo - 1
    For j = myXlonBgnCol To myXlonEndCol
        n = n + 1
        m = Lo - 1
        For i = myXlonBgnRow To myXlonEndRow
            m = m + 1
            myZstrSrsData(m, n) = myZstrOrgData(i, j)
        Next i
    Next j
ExitPath:
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上の連続するデータ範囲を行列で取得する
'Private Sub instCSeriesAry()
'  Dim Lc As Long
'  Dim myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsData() As String
'    'myZstrSrsData(i, j) : 取得文字列
''//クラス内変数への入力
'  Dim myXlonBgnRow As Long: myXlonBgnRow = 8
'  Dim myXlonBgnCol As Long: myXlonBgnCol = 2
'  Dim myXobjSrchSheet As Object: Set myXobjSrchSheet = ActiveSheet
''//CSeriesAryクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsSeriesData As CSeriesAry: Set myXinsSeriesData = New CSeriesAry
'    With myXinsSeriesData
'    '//クラス内変数への入力
'        Set .setSrchSheet = myXobjSrchSheet
'        .letBgnRowCol(1) = myXlonBgnRow
'        .letBgnRowCol(2) = myXlonBgnCol
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonSrsRowCnt = .getSrsRowCnt
'        myXlonSrsColCnt = .getSrsColCnt
'        If myXlonSrsRowCnt <= 0 Or myXlonSrsColCnt <= 0 Then GoTo JumpPath
'        i = myXlonSrsRowCnt + Lo - 1: j = myXlonSrsColCnt + Lo - 1
'        ReDim myZstrSrsData(i, j) As String
'        Lc = .getOptnBase
'        For j = 1 To myXlonSrsColCnt
'            For i = 1 To myXlonSrsRowCnt
'                myZstrSrsData(i + Lo - 1, j + Lo - 1) _
'                    = .getSrsDataAry(i + Lc - 1, j + Lc - 1)
'            Next i
'        Next j
'    End With
'JumpPath:
'    Set myXinsSeriesData = Nothing
'    Set myXobjSrchSheet = Nothing
'    Call variablesOfCSeriesAry(myXlonSrsRowCnt, myZstrSrsData)   'Debug.Print
'End Sub
'Private Sub variablesOfCSeriesAry( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSeriesAryクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
