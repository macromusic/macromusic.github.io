Attribute VB_Name = "xRefVrblToTxt"
'Includes CVrblToTxt
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_変数情報をテキストファイルに書き出す
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefVrblToTxt"
  Private Const meMlonExeNum As Long = 0
 
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//入力データ
  Private myZvarVrbl As Variant, myXstrSpltChar As String, myXstrSaveFilePath As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
  Private myXbisMsgBoxON As Boolean
  
'//モジュール内変数_データ
  Private myZvarPstVrbl As Variant

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisMsgBoxON = False
    myZvarPstVrbl = Empty
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
    Call callxRefVrblToTxt
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            ByVal myZvarVrblIN As Variant, _
            ByVal myXstrSpltCharIN As String, ByVal myXstrSaveFilePathIN As String)
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
    
'//入力変数を初期化
    myZvarVrbl = Empty
    myXstrSpltChar = Empty
    myXstrSaveFilePath = Empty

'//入力変数を取り込み
    myZvarVrbl = myZvarVrblIN
    myXstrSpltChar = myXstrSpltCharIN
    myXstrSaveFilePath = myXstrSaveFilePathIN
    
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
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:変数情報をテキストファイルに書き出す
    Call instCVrblToTxt
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
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

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()

    myXbisMsgBoxON = True
    'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
    'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
    
    ReDim myZvarVrbl(2, 2) As Variant
    myZvarVrbl(1, 1) = "A"
    myZvarVrbl(1, 2) = "A"
    myZvarVrbl(2, 1) = "A"
    myZvarVrbl(2, 2) = "A"
    
    myXstrSpltChar = ""
    
  Dim myXstrPrntPath As String, myXstrFileName As String
    myXstrPrntPath = ThisWorkbook.Path
    myXstrFileName = "testOUT.txt"
    myXstrSaveFilePath = myXstrPrntPath & "\" & myXstrFileName
    
End Sub

'SnsP_
Private Sub snsProc()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_
Private Sub prsProc()
    myXbisExitFlag = False
    
    On Error GoTo ExitPath
    myZvarPstVrbl = myZvarVrbl
    On Error GoTo 0
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
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

'◆ClassProc名_変数情報をテキストファイルに書き出す
Private Sub instCVrblToTxt()
  Dim myXbisCmpltFlag As Boolean
  Dim myXinsVrblToTxt As CVrblToTxt: Set myXinsVrblToTxt = New CVrblToTxt
    With myXinsVrblToTxt
    '//クラス内変数への入力
        .letVrbl = myZvarVrbl
        .letSpltChar = myXstrSpltChar
        .letSaveFilePath = myXstrSaveFilePath
        .letMsgBoxON = True
    '//クラス内プロシージャの実行とクラス内変数からの出力
        myXbisExitFlag = Not .fncbisCmpltFlag
    End With
    Set myXinsVrblToTxt = Nothing
End Sub

'===============================================================================================

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

''SetP_制御用変数を設定する
'Private Sub setControlVariables()
'    myXbisMsgBoxON = True
'    'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
'    'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
'    ReDim myZvarVrbl(2, 2) As Variant
'    myZvarVrbl(1, 1) = "A"
'    myZvarVrbl(1, 2) = "A"
'    myZvarVrbl(2, 1) = "A"
'    myZvarVrbl(2, 2) = "A"
'    myXstrSpltChar = ""
'  Dim myXstrPrntPath As String, myXstrFileName As String
'    myXstrPrntPath = ThisWorkbook.Path
'    myXstrFileName = "testOUT.txt"
'    myXstrSaveFilePath = myXstrPrntPath & "\" & myXstrFileName
'End Sub
'◆ModuleProc名_変数情報をテキストファイルに書き出す
Private Sub callxRefVrblToTxt()
'  Dim myZvarVrbl As Variant, myXstrSpltChar As String, myXstrSaveFilePath As String
'  Dim myXbisCmpltFlag As Boolean
    Call xRefVrblToTxt.callProc(myXbisCmpltFlag, myZvarVrbl, myXstrSpltChar, myXstrSaveFilePath)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefVrblToTxt()
'//xRefVrblToTxtモジュールのモジュールメモリのリセット処理
    Call xRefVrblToTxt.resetConstant
End Sub
