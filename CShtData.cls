VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShtData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfixGetSheetRangeData
'Includes PfixGetSheetRangeDataVariantArray
'Includes PfixGetSheetRangeCommentText

Option Explicit
Option Base 1

'◆ClassProc名_シート上の指定範囲またはデータ全範囲のデータとコメントを取得する
'Rev.001

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CShtData"

'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonDataRowCnt As Long, myXlonDataColCnt As Long, _
            myZstrShtData() As String, myZvarShtData() As Variant, _
            myZstrCmntData() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean

'//PropertySetLet_入力データ
  Private myXobjSheet As Object, myXobjFrstCell As Object, myXobjLastCell As Object

'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getDataRowCnt() As Long
    getDataRowCnt = myXlonDataRowCnt
End Property

Public Property Get getDataColCnt() As Long
    getDataColCnt = myXlonDataColCnt
End Property

Public Property Get getStrShtDataAry(ByVal i As Long, ByVal j As Long) As String
    getStrShtDataAry = myZstrShtData(i, j)
'myZstrShtData(i, j) : 取得文字列
End Property

Public Property Get getVarShtDataAry(ByVal i As Long, ByVal j As Long) As Variant
    getVarShtDataAry = myZvarShtData(i, j)
'myZvarShtData(i, j) : 取得文字列
End Property

Public Property Get getCmntDataAry(ByVal i As Long, ByVal j As Long) As String
    getCmntDataAry = myZstrCmntData(i, j)
'myZstrCmntData(i, j) : 取得コメント
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonDataRowCnt = Empty: myXlonDataColCnt = Empty
    Erase myZstrShtData: Erase myZvarShtData: Erase myZstrCmntData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Set setSheet(ByVal RHS As Object)
    Set myXobjSheet = RHS
End Property

Public Property Let letByVrnt(ByVal RHS As Boolean)
    myXbisByVrnt = RHS
'myXbisByVrnt = False : シートデータをStringで取得する
'myXbisByVrnt = True  : シートデータをVariantで取得する
End Property

Public Property Let letGetCmnt(ByVal RHS As Boolean)
    myXbisGetCmnt = RHS
'myXbisGetCmnt = False : コメントを取得しない
'myXbisGetCmnt = True  : コメントを取得する
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXobjSheet Is Nothing Then GoTo ExitPath
    
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

    If myXbisByVrnt = False Then
    '//シート上の指定範囲またはデータ全範囲のデータを配列に取込む
        Call PfixGetSheetRangeData( _
                myXlonDataRowCnt, myXlonDataColCnt, myZstrShtData, _
                myXobjSheet, myXobjFrstCell, myXobjLastCell)
        
    Else
    '//シート上の指定範囲またはデータ全範囲のデータをVariant配列に取込む
        Call PfixGetSheetRangeDataVariantArray( _
                myXlonDataRowCnt, myXlonDataColCnt, myZvarShtData, _
                myXobjSheet, myXobjFrstCell, myXobjLastCell)
        
    End If
    If myXlonDataRowCnt <= 0 Or myXlonDataRowCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag

    If myXbisGetCmnt = True Then
    '//シート上の指定範囲またはデータ全範囲のコメントを配列に取込む
        Call PfixGetSheetRangeCommentText( _
                myXlonDataRowCnt, myXlonDataColCnt, myZstrCmntData, _
                myXobjSheet, myXobjFrstCell, myXobjLastCell)
    End If
    If myXlonDataRowCnt <= 0 Or myXlonDataRowCnt <= 0 Then GoTo ExitPath
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

 '定型Ｐ_シート上の指定範囲またはデータ全範囲のデータをVariant配列に取込む
Private Sub PfixGetSheetRangeDataVariantArray( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData() As Variant, _
            ByVal myXobjSheet As Object, _
            ByVal myXobjFrstCell As Object, ByVal myXobjLastCell As Object)
'myZvarShtData(i, j) : データ
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZvarShtData
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
    ReDim myZvarShtData(i, j) As Variant
    myZvarRngData = myXobjShtRng.Value
    On Error Resume Next
    If myXlonRowCnt * myXlonColCnt = 1 Then
        myZvarShtData(L, L) = myZvarRngData
    Else
        For j = LBound(myZvarRngData, 2) To UBound(myZvarRngData, 2)
            For i = LBound(myZvarRngData, 1) To UBound(myZvarRngData, 1)
                myZvarShtData(i, j) = myZvarRngData(i, j)
            Next i
        Next j
    End If
    On Error GoTo 0
    Set myXobjShtRng = Nothing: myZvarRngData = Empty
End Sub

 '定型Ｐ_シート上の指定範囲またはデータ全範囲のコメントを配列に取込む
Private Sub PfixGetSheetRangeCommentText( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZstrCmntData() As String, _
            ByVal myXobjSheet As Object, _
            ByVal myXobjFrstCell As Object, ByVal myXobjLastCell As Object)
'myZstrCmntData(i, j) : コメントデータ
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZstrCmntData
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
  Dim myXobjTmpCell As Object, myXstrTmpCmnt As String, i As Long, j As Long
    i = myXlonRowCnt + L - 1: j = myXlonColCnt + L - 1
    ReDim myZstrCmntData(i, j) As String
    On Error Resume Next
    For j = 1 To myXlonColCnt
        For i = 1 To myXlonRowCnt
            Set myXobjTmpCell = myXobjShtRng.Cells(i, j)
            If TypeName(myXobjTmpCell.Comment) = "Comment" Then
                myXstrTmpCmnt = myXobjTmpCell.Comment.Text
            Else
                myXstrTmpCmnt = ""
            End If
            myZstrCmntData(i + L - 1, j + L - 1) = myXstrTmpCmnt
        Next i
    Next j
    On Error GoTo 0
    Set myXobjShtRng = Nothing: Set myXobjTmpCell = Nothing
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_シート上の指定範囲またはデータ全範囲のデータとコメントを取得する
'Private Sub instCShtData()
'  Dim Lc As Long
'  Dim myXlonDataRowCnt As Long, myXlonDataColCnt As Long, _
'            myZstrShtData() As String, myZvarShtData() As Variant, _
'            myZstrCmntData() As String
'    'myZstrShtData(i, j) : 取得文字列
'    'myZvarShtData(i, j) : 取得文字列
'    'myZstrCmntData(i, j) : 取得コメント
''//クラス内変数への入力
'  Dim myXbisByVrnt As Boolean: myXbisByVrnt = False
'    'myXbisByVrnt = False : シートデータをStringで取得する
'    'myXbisByVrnt = True  : シートデータをVariantで取得する
'  Dim myXbisGetCmnt As Boolean: myXbisGetCmnt = False
'    'myXbisGetCmnt = False : コメントを取得しない
'    'myXbisGetCmnt = True  : コメントを取得する
'  Dim myXobjSheet As Object: Set myXobjSheet = ActiveSheet
''//CShtDataクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsShtData As CShtData: Set myXinsShtData = New CShtData
'    With myXinsShtData
'    '//クラス内変数への入力
'        .letByVrnt = myXbisByVrnt
'        .letGetCmnt = myXbisGetCmnt
'        Set .setSheet = myXobjSheet
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonDataRowCnt = .getDataRowCnt
'        myXlonDataColCnt = .getDataColCnt
'        If myXlonDataRowCnt <= 0 Or myXlonDataColCnt <= 0 Then GoTo JumpPath
'        i = myXlonDataRowCnt + Lo - 1: j = myXlonDataColCnt + Lo - 1
'        ReDim myZstrShtData(i, j) As String
'        ReDim myZvarShtData(i, j) As Variant
'        ReDim myZstrCmntData(i, j) As String
'        Lc = .getOptnBase
'        Lc = .getOptnBase
'        If myXbisByVrnt = False Then
'            For j = 1 To myXlonDataColCnt
'                For i = 1 To myXlonDataRowCnt
'                    myZstrShtData(i + Lo - 1, j + Lo - 1) _
'                        = .getStrShtDataAry(i + Lc - 1, j + Lc - 1)
'                Next i
'            Next j
'        Else
'            For j = 1 To myXlonDataColCnt
'                For i = 1 To myXlonDataRowCnt
'                    myZvarShtData(i + Lo - 1, j + Lo - 1) _
'                        = .getVarShtDataAry(i + Lc - 1, j + Lc - 1)
'                Next i
'            Next j
'        End If
'        If myXbisGetCmnt = True Then
'            For j = 1 To myXlonDataColCnt
'                For i = 1 To myXlonDataRowCnt
'                    myZstrCmntData(i + Lo - 1, j + Lo - 1) _
'                        = .getCmntDataAry(i + Lc - 1, j + Lc - 1)
'                Next i
'            Next j
'        End If
'    End With
'JumpPath:
'    Set myXinsShtData = Nothing
'    Set myXobjSheet = Nothing
'    Call variablesOfCShtData(myXlonDataRowCnt, myXlonDataColCnt, myZstrShtData)  'Debug.Print
''    Call variablesOfCShtData(myXlonDataRowCnt, myXlonDataColCnt, myZvarShtData)  'Debug.Print
'    Call variablesOfCShtData(myXlonDataRowCnt, myXlonDataColCnt, myZstrCmntData) 'Debug.Print
'End Sub
'Private Sub variablesOfCShtData( _
'            ByVal myXlonRowCnt As Long, ByVal myXlonColCnt As Long, _
'            ByRef myZvarDataAry As Variant)
''//CSlctShtSrsDataクラス内から出力した変数の内容確認
'    Debug.Print "データ行数: " & myXlonRowCnt
'    Debug.Print "データ列数: " & myXlonColCnt
'    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarDataAry, 1) To UBound(myZvarDataAry, 1)
'        For j = LBound(myZvarDataAry, 2) To UBound(myZvarDataAry, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarDataAry(i, j)
'        Next j
'    Next i
'End Sub
