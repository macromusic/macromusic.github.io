Attribute VB_Name = "xRefShtDataByEnm"
'Includes PfixGetSheetRangeDataVariant
'Includes PfncbisCheckArrayDimension
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上の文字列位置を列挙体で指定してデータを取得する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefShtDataByEnm"
  Private Const meMlonExeNum As Long = 0

'//モジュール内定数_列挙体
Public Enum EnumX
'列挙体使用時の表記 : EnumX.rowX

'//[Sheet1]シート上のパラメータ配置を定義
    shtX = 2                        'Sheet1

    rowFldrPth = 4                  '元フォルダパス ：
    rowFileExt = 5                  '元ファイル拡張子 ：
    colData = 3                     'comment'データ列

    rowData = 8                     'comment'データ行
    colFilePth = 2                  '元ファイル一覧
End Enum
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ_1
  Private myXstrFldrPth As String       '
  Private myXstrFileExt As String       '
  
'//出力データ_2
  Private myZstrFilePth() As String     '
  
'//出力データ_3
  Private myXobjDataRng As Object   '
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData As Variant

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRowCnt = Empty: myXlonColCnt = Empty: myZvarShtData = Empty
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_モジュールメモリのリセット
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()
    
'//処理実行
    Call callxRefShtDataByEnm
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXstrFldrPthOUT As String, myXstrFileExtOUT As String, _
            myZstrFilePthOUT() As String, _
            myXobjDataRngOUT As Object)
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    myXstrFldrPthOUT = Empty: myXstrFileExtOUT = Empty
    Erase myZstrFilePthOUT
    Set myXobjDataRngOUT = Nothing
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXstrFldrPthOUT = myXstrFldrPth
    myXstrFileExtOUT = myXstrFileExt
    
  Dim k As Long
    k = UBound(myZstrFilePth)
    ReDim myZstrFilePthOUT(k) As String
    For k = LBound(myZstrFilePth) To UBound(myZstrFilePth)
        myZstrFilePthOUT(k) = myZstrFilePth(k)
    Next k
    
    Set myXobjDataRngOUT = myXobjDataRng

End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:シート上の全データを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:必要な情報を抽出
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    
    myXstrFldrPth = Empty: myXstrFileExt = Empty
    
    Erase myZstrFilePth

    Set myXobjDataRng = Nothing
End Sub

'RemP_モジュールメモリに保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
End Sub

'SnsP_シート上の全データを取得
Private Sub snsProc()
    myXbisExitFlag = False
    
'//シート上の指定範囲またはデータ全範囲のデータをVariant変数に取込む
  Dim i As Long: i = EnumX.shtX
  Dim myXobjSheet As Object, myXobjFrstCell As Object, myXobjLastCell As Object
    Set myXobjSheet = ThisWorkbook.Worksheets(i)
    
    Call PfixGetSheetRangeDataVariant( _
            myXlonRowCnt, myXlonColCnt, myZvarShtData, _
            myXobjSheet, myXobjFrstCell, myXobjLastCell)
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then GoTo ExitPath
    
    Set myXobjSheet = Nothing
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_必要な情報を抽出
Private Sub prsProc()
    myXbisExitFlag = False
    
  Dim i As Long, j As Long

    On Error Resume Next
    
'//1
    j = EnumX.colData
    
    i = EnumX.rowFldrPth
    myXstrFldrPth = CStr(myZvarShtData(i, j))
    
    i = EnumX.rowFileExt
    myXstrFileExt = CStr(myZvarShtData(i, j))
    
'//2
    i = EnumX.rowData
    j = EnumX.colFilePth
    
  Dim myXstrTmp As String, k As Long, n As Long: n = 0
    For k = i To UBound(myZvarShtData, 1)
        myXstrTmp = Empty
        myXstrTmp = CStr(myZvarShtData(k, j))
        If myXstrTmp = "" Then Exit For
        
        n = n + 1: ReDim Preserve myZstrFilePth(n) As String
        myZstrFilePth(n) = myXstrTmp
    Next k
    
'//3
  Dim myXobjSheet As Object
    k = EnumX.shtX
    Set myXobjSheet = ThisWorkbook.Worksheets(k)
    
    i = EnumX.rowData
    j = EnumX.colFilePth
    Set myXobjDataRng = myXobjSheet.Cells(i, j)
    
'  Dim rb As Long, cb As Long, re As Long, ce As Long
'    rb = EnumX.rowFldrPth
'    cb = EnumX.colData
'    re = EnumX.rowFileExt
'    ce = EnumX.colData
'    With myXobjSheet
'        Set myXobjDataRng = .Range(.Cells(rb, cb), .Cells(re, ce))
'    End With
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    If myXstrFldrPth = "" Then GoTo ExitPath
    If myXstrFileExt = "" Then GoTo ExitPath
    
    If PfncbisCheckArrayDimension(myZstrFilePth, 1) = False Then GoTo ExitPath
    
    If myXobjDataRng Is Nothing Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_使用した変数をモジュールメモリに保存する
Private Sub recProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum"
    myZvarM(1, 2) = meMlonExeNum + 1

  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

 '定型Ｐ_シート上の指定範囲またはデータ全範囲のデータをVariant変数に取込む
Private Sub PfixGetSheetRangeDataVariant( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData As Variant, _
            ByVal myXobjSheet As Object, _
            ByVal myXobjFrstCell As Object, ByVal myXobjLastCell As Object)
'myZvarShtData(i, j) : データ
    myXlonRowCnt = Empty: myXlonColCnt = Empty: myZvarShtData = Empty
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
    myZvarShtData = myXobjShtRng.Value
    Set myXobjShtRng = Nothing
End Sub

 '定型Ｆ_配列変数の次元数が指定次元と一致するかをチェックする
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarOrgData As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarOrgData) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
End Function

 '定型Ｐ_モジュール内定数の値を変更する
Private Sub PfixChangeModuleConstValue(myXbisExitFlag As Boolean, _
            ByVal myXstrMdlName As String, ByRef myZvarM() As Variant)
    myXbisExitFlag = False
    If myXstrMdlName = "" Then GoTo ExitPath
  Dim L As Long, myXvarTmp As Variant
    On Error GoTo ExitPath
    L = LBound(myZvarM, 1): myXvarTmp = myZvarM(L, L)
    On Error GoTo 0
  Dim myXlonDclrLines As Long, myXobjCdMdl As Object
    Set myXobjCdMdl = ThisWorkbook.VBProject.VBComponents(myXstrMdlName).CodeModule
    myXlonDclrLines = myXobjCdMdl.CountOfDeclarationLines
    If myXlonDclrLines <= 0 Then GoTo ExitPath
  Dim i As Long, n As Long
  Dim myXstrTmp As String, myXstrSrch As String, myXstrOrg As String, myXstrRplc As String
Application.DisplayAlerts = False
    For i = 1 To myXlonDclrLines
        myXstrTmp = myXobjCdMdl.Lines(i, 1)
        For n = LBound(myZvarM, 1) To UBound(myZvarM, 1)
            myXstrSrch = "Const" & Space(1) & myZvarM(n, L) & Space(1) & "As" & Space(1)
            If InStr(myXstrTmp, myXstrSrch) > 0 Then
                myXstrOrg = Left(myXstrTmp, InStr(myXstrTmp, "=" & Space(1)) + 1)
                myXstrRplc = myXstrOrg & myZvarM(n, L + 1)
                Call myXobjCdMdl.ReplaceLine(i, myXstrRplc)
            End If
        Next n
    Next i
Application.DisplayAlerts = True
    Set myXobjCdMdl = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'DummyＰ_
Private Sub MsubDummy()
End Sub

'===============================================================================================

'◆ModuleProc名_シート上の文字列位置を列挙体で指定してデータを取得する
Private Sub callxRefShtDataByEnm()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXstrFldrPth As String, myXstrFileExt As String
'  Dim myZstrFilePth() As String
'  Dim myXobjDataRng As Object   '
    Call xRefShtDataByEnm.callProc( _
            myXbisCmpltFlag, myXstrFldrPth, myXstrFileExt, myZstrFilePth, myXobjDataRng)
    Debug.Print myXbisCmpltFlag
    Debug.Print myXstrFldrPth
  Dim k As Long
    For k = LBound(myZstrFilePth) To UBound(myZstrFilePth)
        Debug.Print "データ" & k & ": " & myZstrFilePth(k)
    Next k
    Debug.Print myXobjDataRng.Value
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefShtDataByEnm()
'//xRefShtDataByEnmモジュールのモジュールメモリのリセット処理
    Call xRefShtDataByEnm.resetConstant
End Sub
