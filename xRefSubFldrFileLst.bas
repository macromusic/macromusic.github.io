Attribute VB_Name = "xRefSubFldrFileLst"
'Includes CSubFldrFileLst
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_指定ディレクトリ内のサブファイル一覧を取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSubFldrFileLst"
  Private Const meMlonExeNum As Long = 0
  
'//出力データ
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(k) : ファイルオブジェクト
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
  
'//入力制御信号
  Private myXbisNotOutFileInfo As Boolean
    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
  
'//入力データ
  Private myXstrDirPath As String, myXstrExtsn As String
  Private myXlonSrchOptn As Long
    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
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
    Call callxRefSubFldrFileLst
    
'//処理結果表示
    MsgBox "取得パス数：" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            ByVal myXbisNotOutFileInfoIN As Boolean, _
            ByVal myXstrDirPathIN As String, ByVal myXstrExtsnIN As String, _
            ByVal myXlonSrchOptnIN As Long)
    
'//入力変数を初期化
    myXbisNotOutFileInfo = False
    
    myXstrDirPath = Empty: myXstrExtsn = Empty
    myXlonSrchOptn = Empty

'//入力変数を取り込み
    myXbisNotOutFileInfo = myXbisNotOutFileInfoIN
    
    myXstrDirPath = myXstrDirPathIN
    myXstrExtsn = myXstrExtsnIN
    myXlonSrchOptn = myXlonSrchOptnIN
    
'//出力変数を初期化
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    
'//処理実行
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//出力変数に格納
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    
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
    
'//S:指定ディレクトリ内の複数サブフォルダ内のサブファイル一覧を取得
    Call instCSubFldrFileLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
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
    
'    If myXstrDirPath = "" Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()

    myXstrDirPath = ActiveWorkbook.Path
    
    myXstrExtsn = ""
    
    myXlonSrchOptn = 1
    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
    
    myXbisNotOutFileInfo = False
    'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
    'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない

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

'◆ClassProc名_指定ディレクトリ内の複数サブフォルダ内のサブファイル一覧を取得する
Private Sub instCSubFldrFileLst()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSubFldrFileLst As CSubFldrFileLst
    Set myXinsSubFldrFileLst = New CSubFldrFileLst
    With myXinsSubFldrFileLst
    '//クラス内変数への入力
        .letDirPath = myXstrDirPath
        .letExtsn = myXstrExtsn
        .letSrchOptn = myXlonSrchOptn
        .letNotOutFileInfo = myXbisNotOutFileInfo
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonFileCnt = .getFileCnt
        If myXlonFileCnt <= 0 Then GoTo JumpPath
        k = myXlonFileCnt + Lo - 1
        ReDim myZobjFile(k) As Object
        ReDim myZstrFileName(k) As String
        ReDim myZstrFilePath(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFileCnt
            Set myZobjFile(k + Lo - 1) = .getFileAry(k + Lc - 1)
            myZstrFileName(k + Lo - 1) = .getFileNameAry(k + Lc - 1)
            myZstrFilePath(k + Lo - 1) = .getFilePathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSubFldrFileLst = Nothing
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
'    myXstrDirPath = ActiveWorkbook.Path
'    myXstrExtsn = ""
'    myXlonSrchOptn = 1
'    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
'    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
'    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
'    myXbisNotOutFileInfo = False
'    'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
'    'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
'End Sub
'◆ModuleProc名_指定ディレクトリ内のサブファイル一覧を取得する
Private Sub callxRefSubFldrFileLst()
'  Dim myXbisNotOutFileInfo As Boolean, _
'        myXstrDirPath As String, myXstrExtsn As String, myXlonSrchOptn As Long
'    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
'    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
'    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
'    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
'    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String
'    'myZobjFile(k) : ファイルオブジェクト
'    'myZstrFileName(k) : ファイル名
'    'myZstrFilePath(k) : ファイルパス
    Call xRefSubFldrFileLst.callProc( _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXbisNotOutFileInfo, myXstrDirPath, myXstrExtsn, myXlonSrchOptn)
    Debug.Print "データ: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrFileLst()
'//xRefSubFldrFileLstモジュールのモジュールメモリのリセット処理
    Call xRefSubFldrFileLst.resetConstant
End Sub
