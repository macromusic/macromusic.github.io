Attribute VB_Name = "xRefRunFilesDblPlus"
'Includes m1MexeFileLstup
'Includes m1MexeFileLstupSub
'Includes xRefRunFilesDbl
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_複数ファイルをリストアップして連続処理を実施する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRunFilesDblPlus"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数

'//モジュール内定数_列挙体
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  
'//入力制御信号
  
'//入力データ
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Dim myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
  Dim myXlonSubFileCnt As Long, myZstrSubFilePath() As String
    'myZstrOrgFilePath(i) : 元ファイルパス1
    'myZstrSubFilePath(i) : 元ファイルパス2
    

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
    myXlonSubFileCnt = Empty: Erase myZstrSubFilePath
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

'//プログラム構成
    '入力: -
    '処理:  '◆ModuleProc名_処理ファイルをリストアップする
            '◆ModuleProc名_2種類の複数ファイルに対して連続処理を実施する
    '出力: -
    
'//処理実行
    Call callxRefRunFilesDblPlus
    
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
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:処理ファイルをリストアップ1
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:処理ファイルをリストアップ2
    Call snsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:2種類の複数ファイルに対して連続処理を実施
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
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
End Sub

'SnsP_処理ファイルをリストアップする1
Private Sub snsProc1()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object, myXstrExtsn As String
    'myZobjFile(k) : ファイルオブジェクト
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
    
    Call m1MexeFileLstup.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgFileCnt = myXlonFileCnt
    myZstrOrgFilePath() = myZstrFilePath()
    
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_処理ファイルをリストアップする1
Private Sub snsProc2()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object, myXstrExtsn As String
    'myZobjFile(k) : ファイルオブジェクト
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
    
    Call m1MexeFileLstupSub.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonSubFileCnt = myXlonFileCnt
    myZstrSubFilePath() = myZstrFilePath()
    
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_2種類の複数ファイルに対して連続処理を実施する
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeOrgFileCnt As Long, _
        myZstrExeOrgFileName() As String, myZstrOrgExeFilePath() As String
    'myZstrExeOrgFileName(i) : 実行ファイル名1
    'myZstrOrgExeFilePath(i) : 実行ファイルパス1
  Dim myXlonExeSubFileCnt As Long, _
        myZstrExeSubFileName() As String, myZstrSubExeFilePath() As String
    'myZstrExeSubFileName(i) : 実行ファイル名2
    'myZstrSubExeFilePath(i) : 実行ファイルパス2
    
    Call xRefRunFilesDbl.callProc( _
            myXbisCompFlag, _
            myXlonExeOrgFileCnt, myZstrExeOrgFileName, myZstrExeOrgFilePath, _
            myXlonExeSubFileCnt, myZstrExeSubFileName, myZstrExeSubFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath, _
            myXlonSubFileCnt, myZstrSubFilePath)
    If myXlonExeOrgFileCnt <= 0 Or myXlonExeSubFileCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExeOrgFileName: Erase myZstrOrgExeFilePath
    Erase myZstrExeSubFileName: Erase myZstrSubExeFilePath
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_
Private Sub runProc()
    myXbisExitFlag = False
    
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

'◆ModuleProc名_複数ファイルをリストアップして連続処理を実施する
Private Sub callxRefRunFilesDblPlus()
  Dim myXbisCompFlag As Boolean
    Call xRefRunFilesDblPlus.callProc(myXbisCompFlag)
    Debug.Print "結果: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesDblPlus()
'//xRefRunFilesDblPlusモジュールのモジュールメモリのリセット処理
    Call xRefRunFilesDblPlus.resetConstant
End Sub
