Attribute VB_Name = "xRefChartOperation"
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルシート内の全グラフに対して処理を実行する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefChartOperation"
  Private Const meMlonExeNum As Long = 0
  Private Const meMvarField As Variant = Empty
  
'//モジュール内定数
  Private Const coXvarField As Variant = ""

'//モジュール内定数_列挙体
Private Enum EnumX
'列挙体使用時の表記 : EnumX.rowX
'■myEnumの表記ルール
    '�@シートNo. : "sht" & "Enum名" & " = " & "値" & "'シート名"
    '�A行No.     : "row" & "Enum名" & " = " & "値" & "'検索するシート上の文字列"
    '�B列No.     : "col" & "Enum名" & " = " & "値" & "'検索するシート上の文字列"
    '�C行No.     : "row" & "Enum名" & " = " & "値" & "'comment" & "'検索するコメントの文字列"
    '�D列No.     : "col" & "Enum名" & " = " & "値" & "'comment" & "'検索するコメントの文字列"
    shtX = 1        'Sheet1
'    rowX = 1        '行No
'    colX = 1        '列No
'    rowY = 1        'comment'行No
'    colY = 1        'comment'列No
End Enum
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjSheet As Object
  Private myXlonChrtObjCnt As Long, myXlonErrChrtObjCnt As Long, myZobjErrChrtObj() As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjSheet = Nothing
    myXlonChrtObjCnt = Empty
    myXlonErrChrtObjCnt = Empty
    Erase myZobjErrChrtObj
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_モジュールメモリのリセット
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
'    myZvarM(2, 1) = "meMvarField": myZvarM(2, 2) = Chr(34) & Chr(34)
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()

'//プログラム構成
    '入力: -
    '処理: -
    '出力: -
    
'//処理実行
    Call callxRefChartOperation
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables

'//エクセルシート内の全グラフに対して処理を実行
    Call PabsChartOperation( _
            myXbisExitFlag, myXlonChrtObjCnt, myXlonErrChrtObjCnt, myZobjErrChrtObj, _
            myXobjSheet)
    If myXbisExitFlag = True Then Exit Sub
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'RemP_モジュールメモリに保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
'    If myXvarFieldIN = Empty Then myXvarFieldIN = meMvarField
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()

    Set myXobjOrg = Selection

End Sub

 '抽象Ｐ_エクセルシート内の全グラフに対して処理を実行する
Private Sub PabsChartOperation( _
            myXbisExitFlag As Boolean, myXlonChrtObjCnt As Long, _
            myXlonErrChrtObjCnt As Long, myZobjErrChrtObj() As Object, _
            ByVal myXobjSheet As Object)
    myXlonChrtObjCnt = Empty: myXlonErrChrtObjCnt = Empty: Erase myZobjErrChrtObj
    On Error GoTo ExitPath
  Dim k As Long: k = myXobjSheet.ChartObjects.Count
    If k <= 0 Then GoTo ExitPath
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjChrtObj As Object, n As Long, e As Long: n = 0: e = Lo - 1
    For Each myXobjChrtObj In myXobjSheet.ChartObjects
        Call PsubChartOperation(myXbisExitFlag, myXobjChrtObj)
        If myXbisExitFlag = True Then
            e = e + 1: ReDim Preserve myZobjErrChrtObj(e) As Object
            Set myZobjErrChrtObj(e) = myXobjChrtObj
        Else
            n = n + 1
        End If
    Next myXobjChrtObj
    myXlonChrtObjCnt = n: myXlonErrChrtObjCnt = e - Lo + 1
    If myXlonErrChrtObjCnt >= 1 Then
        myXbisExitFlag = True
    Else
        myXbisExitFlag = False
    End If
    Set myXobjChrtObj = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub
Private Sub PsubChartOperation(myXbisExitFlag As Boolean, _
            ByVal myXobjChrtObj As Object)
    myXbisExitFlag = False
'//シート内の全図形に対する処理
'    XarbProgCode
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
'    myZvarM(1, 1) = "meMvarField"
'    myZvarM(1, 2) = myXvarField

  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

'モジュール内Ｐ_
Private Sub MsubProc()
End Sub

'モジュール内Ｆ_
Private Function MfncFunc() As Variant
End Function

'===============================================================================================

 '定型Ｐ_
Private Sub PfixProc()
End Sub

 '定型Ｆ_
Private Function PfncFunc() As Variant
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

'◆ModuleProc名_エクセルシート内の全グラフに対して処理を実行する
Private Sub callxRefChartOperation()
  Dim myXbisCompFlag As Boolean
    Call xRefChartOperation.callProc(myXbisCompFlag)
    Debug.Print "完了: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefChartOperation()
'//xRefChartOperationモジュールのモジュールメモリのリセット処理
    Call xRefChartOperation.resetConstant
End Sub
