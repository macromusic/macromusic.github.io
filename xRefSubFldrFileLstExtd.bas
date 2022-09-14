Attribute VB_Name = "xRefSubFldrFileLstExtd"
'Includes CSubFldrFileLst
'Includes CVrblToSht
'Includes PfncstrCheckAndGetFilesParentFolder
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_指定ディレクトリ内のサブファイル一覧を取得してシートに書き出す
'Rev.003
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSubFldrFileLstExtd"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(k) : ファイルオブジェクト
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
  Private myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
  
'//入力制御信号
  Private myXbisNotOutFileInfo As Boolean
    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す
  
'//入力データ
  Private myXstrDirPath As String, myXstrExtsn As String
    'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
    'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
  Private myXlonSrchOptn As Long
    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
  Private myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean, myXbisPstFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonFileOrgCnt As Long, myZstrFilePathOrg() As String
  Private myZvarPstData As Variant, myXobjPstFrstCell As Object
  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myXobjPstdCell As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False: myXbisPstFlag = False
    
    myXlonFileOrgCnt = Empty: Erase myZstrFilePathOrg
    myZvarPstData = Empty: Set myXobjPstFrstCell = Nothing
    myXbisInptBxOFF = False: myXbisEachWrtON = False
    Set myXobjPstdCell = Nothing
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
    Call callxRefSubFldrFileLstExtd
    
'//処理結果表示
    MsgBox "取得パス数：" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, myXobjDirPstdCellOUT As Object, _
            ByVal myXbisNotOutFileInfoIN As Boolean, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrDirPathIN As String, ByVal myXstrExtsnIN As String, _
            ByVal myXlonSrchOptnIN As Long, _
            ByVal myXobjDirPstFrstCellIN As Object, ByVal myXobjFilePstFrstCellIN As Object)
    
'//入力変数を初期化
    myXbisNotOutFileInfo = False
    myXlonOutputOptn = Empty
    
    myXstrDirPath = Empty: myXstrExtsn = Empty
    myXlonSrchOptn = Empty
    Set myXobjDirPstFrstCell = Nothing: Set myXobjFilePstFrstCell = Nothing

'//入力変数を取り込み
    myXbisNotOutFileInfo = myXbisNotOutFileInfoIN
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrDirPath = myXstrDirPathIN
    myXstrExtsn = myXstrExtsnIN
    myXlonSrchOptn = myXlonSrchOptnIN
    Set myXobjDirPstFrstCell = myXobjDirPstFrstCellIN
    Set myXobjFilePstFrstCell = myXobjFilePstFrstCellIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    
'//処理実行
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    Set myXobjFilePstdCellOUT = myXobjFilePstdCell
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables1
    Call setControlVariables2
    
'//S:指定ディレクトリ内の複数サブフォルダ内のサブファイル一覧を取得
    Call instCSubFldrFileLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:ファイルパスをシートに書き出す
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
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
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
Private Sub setControlVariables1()

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

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    myXlonOutputOptn = 3
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す

'    myZvarVrbl = 1
    
'    Set myXobjDirPstFrstCell = Selection
'    Set myXobjFilePstFrstCell = Selection
    
    myXbisInptBxOFF = False
    'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
    'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
    
    myXbisEachWrtON = False
    'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
    'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
    
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

'RunP_ファイルパスをシートに書き出す
Private Sub runProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt1 As String _
        = "ファイルパスを貼り付ける位置を指定して下さい。"
  Const coXstrMsgBxPrmpt2 As String _
        = "ファイル名を貼り付ける位置を指定して下さい。"
  Const coXstrMsgBxPrmpt3 As String _
        = "ディレクトリパスを貼り付ける位置を指定して下さい。"
   
    If myXlonOutputOptn = 0 Then Exit Sub
    
'//ファイルパス一覧の親フォルダが同一か確認して同一であれば親フォルダパスを取得
  Dim myXstrPrntPath As String
    myXstrPrntPath = PfncstrCheckAndGetFilesParentFolder(myZstrFilePath)
    If myXstrPrntPath = "" Then myXlonOutputOptn = 1
        
'//ファイルパスをシートに書き出す方法で分岐
    If myXlonOutputOptn = 2 Then
    '//ファイル名を書き出す場合
        myZvarPstData = myZstrFileName
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt2
        
    ElseIf myXlonOutputOptn = 3 Then
    '//親フォルダに応じて書き出す場合
    
    '//ディレクトリパスをエクセルシートに書き出す
        myZvarPstData = myXstrPrntPath
        Set myXobjPstFrstCell = myXobjDirPstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt3
        
        Call instCVrblToSht
        If myXbisPstFlag = False Then GoTo ExitPath
        Set myXobjDirPstdCell = myXobjPstdCell
        
    '//ファイル名をエクセルシートに書き出す
        myZvarPstData = myZstrFileName
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt2
        
    Else
    '//ファイルパスを書き出す場合
        myZvarPstData = myZstrFilePath
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt1
            
    End If
    
'//ファイルパスをエクセルシートに書き出す
    Call instCVrblToSht
    If myXbisPstFlag = False Then GoTo ExitPath
    Set myXobjFilePstdCell = myXobjPstdCell
    
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

'◆ClassProc名_変数情報をエクセルシートに書き出す
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//クラス内変数への入力
        .letVrbl = myZvarPstData
        Set .setPstFrstCell = myXobjPstFrstCell
        .letInptBxOFF = myXbisInptBxOFF
        .letEachWrtON = myXbisEachWrtON
    '//クラス内プロシージャの実行とクラス内変数からの出力
        myXbisPstFlag = .fncbisCmpltFlag
        Set myXobjPstdCell = .getPstdRng
    End With
    Set myXinsVrblToSht = Nothing
End Sub

'===============================================================================================

 '定型Ｆ_ファイルパス一覧の親フォルダが同一か確認して同一であれば親フォルダパスを取得する
Private Function PfncstrCheckAndGetFilesParentFolder( _
            ByRef myZstrOrgFilePath() As String) As String
    PfncstrCheckAndGetFilesParentFolder = Empty
'//ファイルの親フォルダを取得
  Dim myXstrTmpFile As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgFilePath): myXstrTmpFile = myZstrOrgFilePath(L)
    On Error GoTo 0
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrPrntPath As String
    myXstrPrntPath = myXobjFSO.GetParentFolderName(myXstrTmpFile)
'//全ファイルの親フォルダが同一か確認
  Dim myXvarTmp As Variant, myXstrTmp As String
    For Each myXvarTmp In myZstrOrgFilePath
        myXstrTmp = myXobjFSO.GetParentFolderName(myXvarTmp)
        If myXstrPrntPath <> myXstrTmp Then GoTo ExitPath
    Next myXvarTmp
    PfncstrCheckAndGetFilesParentFolder = myXstrPrntPath
ExitPath:
    Set myXobjFSO = Nothing
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

''SetP_制御用変数を設定する
'Private Sub setControlVariables1()
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
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    myXlonOutputOptn = 3
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
'    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
'    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す
''    myZvarVrbl = 1
''    Set myXobjDirPstFrstCell = Selection
''    Set myXobjFilePstFrstCell = Selection
'End Sub
'◆ModuleProc名_指定ディレクトリ内のサブファイル一覧を取得してシートに書き出す
Private Sub callxRefSubFldrFileLstExtd()
'  Dim myXbisNotOutFileInfo As Boolean, myXlonOutputOptn As Long, _
'        myXstrDirPath As String, myXstrExtsn As String, myXlonSrchOptn As Long, _
'        myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
'    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
'    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
'    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
'    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す
'    'myXlonSrchOptn = 1 : 指定フォルダ直下のファイルのパスのみ取得
'    'myXlonSrchOptn = 2 : 指定フォルダ直下のファイルとサブフォルダ内のファイルのパスを取得
'    'myXlonSrchOptn = 3 : 指定フォルダ直下のサブフォルダ内のファイルのパスのみ取得
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
'    'myZobjFile(k) : ファイルオブジェクト
'    'myZstrFileName(k) : ファイル名
'    'myZstrFilePath(k) : ファイルパス
    Call xRefSubFldrFileLstExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXbisNotOutFileInfo, myXlonOutputOptn, _
            myXstrDirPath, myXstrExtsn, myXlonSrchOptn, _
            myXobjDirPstFrstCell, myXobjFilePstFrstCell)
    Debug.Print "データ: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrFileLstExtd()
'//xRefSubFldrFileLstExtdモジュールのモジュールメモリのリセット処理
    Call xRefSubFldrFileLstExtd.resetConstant
End Sub
