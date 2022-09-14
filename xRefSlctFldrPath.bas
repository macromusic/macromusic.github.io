Attribute VB_Name = "xRefSlctFldrPath"
'Includes CSlctFldrPath
'Includes CExplrAdrs
'Includes CExplrAdrsSlct
'Includes PfncbisCheckFolderExist
'Includes PfncobjGetFolder
'Includes PfixGetFolderNameInformationByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_フォルダを選択してそのパスを取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSlctFldrPath"
  Private Const meMlonExeNum As Long = 0
  
'//出力データ
  Private myXstrFldrPath As String, myXobjFldr As Object, _
            myXstrDirPath As String, myXstrFldrName As String
  
'//入力制御信号
  Private myXlonDirSlctOptn As Long
    'myXlonDirSlctOptn = 1 : アクティブブックの親フォルダパスを取得
    'myXlonDirSlctOptn = 2 : FileDialogオブジェクトを使用してフォルダパスを取得
    'myXlonDirSlctOptn = 3 : 最前面のエクスプローラに表示されているフォルダパスを取得
    'myXlonDirSlctOptn = 4 : 起動中のエクスプローラを選択してそのアドレスバーを取得
  
'//入力データ
  Private myXstrDfltFldrPath As String
    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
  Private myXlonIniView As Long
    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
  Private myXbisExplrAdrsMsgOptn As Boolean
    'myXbisExplrAdrsMsgOptn = True  : ウィンド選択の確認メッセージを表示する
    'myXbisExplrAdrsMsgOptn = False : ウィンド選択の確認メッセージを表示しない
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  Private myXbisCurDirON As Boolean
    'myXbisCurDirON = False : デフォルトパスにカレントディレクトリを設定しない
    'myXbisCurDirON = True  : デフォルトパスにカレントディレクトリを設定する

'//モジュール内変数_データ
    
'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    myXbisCurDirON = False
    
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
    Call callxRefSlctFldrPath
    
'//処理結果表示
    MsgBox "取得パス：" & myXstrFldrPath
    
End Sub

'PublicP_
Public Sub callProc( _
            myXstrFldrPathOUT As String, myXobjFldrOUT As Object, _
            myXstrDirPathOUT As String, myXstrFldrNameOUT As String, _
            ByVal myXlonDirSlctOptnIN As Long, _
            ByVal myXstrDfltFldrPathIN As String, ByVal myXlonIniViewIN As Long, _
            ByVal myXbisExplrAdrsMsgOptnIN As Boolean)

'//入力変数を初期化
    myXlonDirSlctOptn = Empty
    myXstrDfltFldrPath = Empty: myXlonIniView = Empty
    myXbisExplrAdrsMsgOptn = False

'//入力変数を取り込み
    myXlonDirSlctOptn = myXlonDirSlctOptnIN
    myXstrDfltFldrPath = myXstrDfltFldrPathIN
    myXlonIniView = myXlonIniViewIN
    myXbisExplrAdrsMsgOptn = myXbisExplrAdrsMsgOptnIN
    
'//出力変数を初期化
    myXstrFldrPathOUT = Empty: Set myXobjFldrOUT = Nothing
    myXstrDirPathOUT = Empty: myXstrFldrNameOUT = Empty
    
'//処理実行
    Call ctrProc
    If myXstrFldrPath = "" Then Exit Sub
    
'//出力変数に格納
    myXstrFldrPathOUT = myXstrFldrPath
    Set myXobjFldrOUT = myXobjFldr
    myXstrDirPathOUT = myXstrDirPath
    myXstrFldrNameOUT = myXstrFldrName

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
    
'//S:フォルダパスを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
        
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXstrFldrPath = False
    myXstrFldrPath = Empty: Set myXobjFldr = Nothing
    myXstrDirPath = Empty: myXstrFldrName = Empty
End Sub

'RemP_モジュールメモリに保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
    If meMlonExeNum > 0 Then myXbisCurDirON = True
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'    If myXlonDirSlctOptn < 1 Or myXlonDirSlctOptn > 4 Then myXlonDirSlctOptn = 2
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
        
    myXlonDirSlctOptn = 2
    'myXlonDirSlctOptn = 1 : アクティブブックの親フォルダパスを取得
    'myXlonDirSlctOptn = 2 : FileDialogオブジェクトを使用してフォルダパスを取得
    'myXlonDirSlctOptn = 3 : 最前面のエクスプローラに表示されているフォルダパスを取得
    'myXlonDirSlctOptn = 4 : 起動中のエクスプローラを選択してそのアドレスバーを取得
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    If myXbisCurDirON = True Then myXstrDfltFldrPath = CurDir
    
  Dim myXbisTmp As Boolean
    If myXstrDfltFldrPath = "" Or myXstrDfltFldrPath = "C" Or _
            myXstrDfltFldrPath = "1" Or myXstrDfltFldrPath = "2" Then
        myXbisTmp = PfncbisCheckFolderExist(myXstrDfltFldrPath)
        If myXbisTmp = False Then myXstrDfltFldrPath = "2"
    End If
    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
    
    myXlonIniView = msoFileDialogViewList
    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables3()
    
    myXbisExplrAdrsMsgOptn = True
    'myXbisExplrAdrsMsgOptn = True  : ウィンド選択の確認メッセージを表示する
    'myXbisExplrAdrsMsgOptn = False : ウィンド選択の確認メッセージを表示しない
    
End Sub

'SnsP_フォルダパスを取得
Private Sub snsProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt As String _
        = "ダイアログボックスが表示されますので、フォルダを選択して下さい。"
    
'//フォルダパスの取得方法で分岐してパスを取得
    Select Case myXlonDirSlctOptn
        Case 1
        '//アクティブブックの親フォルダを取得
            myXstrFldrPath = ActiveWorkbook.Path
            
        Case 2
        '//FileDialogオブジェクトを使用してフォルダを選択
            Call setControlVariables2
            MsgBox coXstrMsgBxPrmpt
            Call instCSlctFldrPath
            
        Case 3
        '//CExplrAdrsインスタンスを使用してフォルダを取得
            Call setControlVariables3
            Call instCExplrAdrs
            
        Case 4
        '//CExplrAdrsSlctインスタンスを使用してフォルダを取得
            Call instCExplrAdrsSlct
            
        Case Else
    End Select
    If myXstrFldrPath = "" Then GoTo ExitPath
    
'//指定フォルダの存在を確認
    If PfncbisCheckFolderExist(myXstrFldrPath) = False Then
        myXstrFldrPath = ""
        GoTo ExitPath
    End If
    
'//指定フォルダのオブジェクトを取得
    Set myXobjFldr = PfncobjGetFolder(myXstrFldrPath)
    
'//指定フォルダのフォルダ名情報を取得
    Call PfixGetFolderNameInformationByFSO(myXstrDirPath, myXstrFldrName, myXstrFldrPath)
    
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

'◆ClassProc名_フォルダを選択してそのパスを取得する
Private Sub instCSlctFldrPath()
  Dim myXinsFldrPath As CSlctFldrPath: Set myXinsFldrPath = New CSlctFldrPath
    With myXinsFldrPath
    '//クラス内変数への入力
        .letDfltFldrPath = myXstrDfltFldrPath
        .letIniView = myXlonIniView
    '//クラス内プロシージャの実行とクラス内変数からの出力
        myXstrFldrPath = .fncstrDirectoryPath
    End With
    Set myXinsFldrPath = Nothing
End Sub

'◆ClassProc名_起動中のエクスプローラのアドレスバーを取得する
Private Sub instCExplrAdrs()
  Dim myXinsExplrAdrs As CExplrAdrs: Set myXinsExplrAdrs = New CExplrAdrs
    With myXinsExplrAdrs
        .letMsgOptn = myXbisExplrAdrsMsgOptn
        myXstrFldrPath = .fncstrExplorerAddress
    End With
    Set myXinsExplrAdrs = Nothing
End Sub

'◆ClassProc名_起動中のエクスプローラを選択してそのアドレスバーを取得する
Private Sub instCExplrAdrsSlct()
  Dim myXinsExplrAdrsSlct As CExplrAdrsSlct: Set myXinsExplrAdrsSlct = New CExplrAdrsSlct
    With myXinsExplrAdrsSlct
        myXstrFldrPath = .fncstrExplorerAddress
    End With
    Set myXinsExplrAdrsSlct = Nothing
End Sub

'===============================================================================================

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｆ_指定フォルダのオブジェクトを取得する
Private Function PfncobjGetFolder(ByVal myXstrDirPath As String) As Object
    Set PfncobjGetFolder = Nothing
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FolderExists(myXstrDirPath) = False Then Exit Function
        Set PfncobjGetFolder = .GetFolder(myXstrDirPath)
    End With
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_指定フォルダのフォルダ名情報を取得する(FileSystemObject使用)
Private Sub PfixGetFolderNameInformationByFSO( _
            myXstrPrntPath As String, myXstrDirName As String, _
            ByVal myXstrDirPath As String)
    myXstrPrntPath = Empty: myXstrDirName = Empty
    If myXstrDirPath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        myXstrPrntPath = .GetParentFolderName(myXstrDirPath)    '親フォルダパス
        myXstrDirName = .GetFolder(myXstrDirPath).Name          'フォルダ名
    End With
    Set myXobjFSO = Nothing
End Sub

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
'    myXlonDirSlctOptn = 2
'    'myXlonDirSlctOptn = 1 : アクティブブックの親フォルダパスを取得
'    'myXlonDirSlctOptn = 2 : FileDialogオブジェクトを使用してフォルダパスを取得
'    'myXlonDirSlctOptn = 3 : 最前面のエクスプローラに表示されているフォルダパスを取得
'    'myXlonDirSlctOptn = 4 : 起動中のエクスプローラを選択してそのアドレスバーを取得
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    If myXbisCurDirON = True Then myXstrDfltFldrPath = CurDir
'  Dim myXstrTmpPath As String
'    If myXstrDfltFldrPath = "" Or myXstrDfltFldrPath = "C" Or _
'            myXstrDfltFldrPath = "1" Or myXstrDfltFldrPath = "2" Then
'        myXstrTmpPath = PfncbisCheckFolderExist(myXstrDfltFldrPath)
'        If myXstrTmpPath = False Then myXstrDfltFldrPath = "2"
'    End If
'    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
'    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
'    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
'    myXlonIniView = msoFileDialogViewList
'    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
'    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
'    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
'    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
'    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
'    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
'    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables3()
'    myXbisExplrAdrsMsgOptn = True
'    'myXbisExplrAdrsMsgOptn = True  : ウィンド選択の確認メッセージを表示する
'    'myXbisExplrAdrsMsgOptn = False : ウィンド選択の確認メッセージを表示しない
'End Sub
'◆ModuleProc名_ディレクトリを選択してそのパスを取得する
Private Sub callxRefSlctFldrPath()
'  Dim myXlonDirSlctOptn As Long, _
'        myXstrDfltFldrPath As String, myXlonIniView As Long, _
'        myXbisExplrAdrsMsgOptn As Boolean
'  Dim myXstrFldrPath As String, myXobjFldr As Object, _
'        myXstrDirPath As String, myXstrFldrName As String
    Call xRefSlctFldrPath.callProc( _
            myXstrFldrPath, myXobjFldr, myXstrDirPath, myXstrFldrName, _
            myXlonDirSlctOptn, myXstrDfltFldrPath, myXlonIniView, myXbisExplrAdrsMsgOptn)
    Debug.Print "データ: " & myXstrFldrPath
    Debug.Print "データ: " & myXstrDirPath
    Debug.Print "データ: " & myXstrFldrName
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSlctFldrPath()
'//xRefSlctFldrPathモジュールのモジュールメモリのリセット処理
    Call xRefSlctFldrPath.resetConstant
End Sub
