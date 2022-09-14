Attribute VB_Name = "xRefRunFilesShtFile"
'Includes xRefGetShtFileLst
'Includes xRefRunFiles
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上のファイルパスを選択して連続処理を実施する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRunFilesShtFile"
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
  Private myXbisByDscrt As Boolean
  Private myXlonRngOptn As Long
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
    'myXbisByDscrt = False : シート上の連続範囲を指定して取得する
    'myXbisByDscrt = True  : シート上の不連続範囲を指定して取得する
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
  
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
    'myZstrOrgFilePath(i) : 元ファイルパス

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisByDscrt = False
    myXlonRngOptn = Empty
    myXstrInptBxPrmpt = Empty: myXstrInptBxTtl = Empty
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
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
    '処理:  '◆ModuleProc名_エクセルシート上に記載されたデータを選択してパス一覧を取得する
            '◆ModuleProc名_複数ファイルに対して連続処理を実施する
    '出力: -
    
'//処理実行
    Call callxRefRunFilesShtFile
    
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
    
'//S:処理ファイルパスを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:複数ファイルに対して連続処理を実施
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:
    Call runProc
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
    myXbisByDscrt = False
    'myXbisByDscrt = False : シート上の連続範囲を指定して取得する
    'myXbisByDscrt = True  : シート上の不連続範囲を指定して取得する
    myXlonRngOptn = 0
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
    myXstrInptBxPrmpt = "処理したいファイルパスを選択して下さい。"
    myXstrInptBxTtl = "ファイルパスの選択"
End Sub

'SnsP_処理ファイルパスを取得する
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
    Call xRefGetShtFileLst.callProc( _
            myXlonFileCnt, myZstrFileName, myZstrFilePath)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgFileCnt = myXlonFileCnt
    myZstrOrgFilePath() = myZstrFilePath()
    
    Erase myZstrFileName: Erase myZstrFilePath
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_複数ファイルに対して連続処理を実施する
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeFileCnt As Long, _
        myZstrExeFileName() As String, myZstrExeFilePath() As String
    'myZstrExeFileName(i) : 実行ファイル名
    'myZstrExeFilePath(i) : 実行ファイルパス
    
    Call xRefRunFiles.callProc( _
            myXbisCompFlag, _
            myXlonExeFileCnt, myZstrExeFileName, myZstrExeFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath)
    If myXlonExeFileCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExeFileName: Erase myZstrExeFilePath
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

'◆ModuleProc名_シート上のファイルパスを選択して連続処理を実施する
Private Sub callxRefRunFilesShtFile()
  Dim myXbisCompFlag As Boolean
    Call xRefRunFilesShtFile.callProc(myXbisCompFlag)
    Debug.Print "結果: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesShtFile()
'//xRefRunFilesShtFileモジュールのモジュールメモリのリセット処理
    Call xRefRunFilesShtFile.resetConstant
End Sub
