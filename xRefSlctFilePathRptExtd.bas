Attribute VB_Name = "xRefSlctFilePathRptExtd"
'Includes CSlctFilePathRpt
'Includes CVrblToSht
'Includes PincPickUpExtensionMatchFilePathArray
'Includes PfncbisCheckFileExtension
'Includes PfixGetFileFor1DArray
'Includes PfixGetFolderFileStringInformationFor1DArray
'Includes PfncstrCheckAndGetFilesParentFolder
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_指定文字列を含むファイル名のファイルを繰返し選択してそのパスを取得してシートに書き出す
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSlctFilePathRptExtd"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
            myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(i) : ファイルオブジェクト
    'myZstrFileName(i) : ファイル名
    'myZstrFilePath(i) : ファイルパス
  Private myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
  
'//入力制御信号
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す
  
'//入力データ
  Private myXstrDfltFldrPath As String
    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
  Private myXstrDfltFilePath As String
    'myXstrDfltFilePath = ""  : デフォルトパス指定無し
    'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
    'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
    'myXstrDfltFilePath = "*" : デフォルトパスを指定
  Private myXstrExtsn As String
  Private myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long
    'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
    'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
    'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
    'myXbisFltrClr = True  : ファイルフィルタを初期化する
    'myXlonFltrIndx = 1〜 : ファイルフィルタの初期選択
  Private myXlonIniView As Long, myXbisMultSlct As Boolean
    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
    'myXbisMultSlct = False : 複数のファイルを選択不可能
    'myXbisMultSlct = True  : 複数のファイルを選択可能
  Private myXlonOrdrCnt As Long, myXlonTrgtWrdCnt As Long, myZvarOdrTrgtWrdPos() As Variant
    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
  Private myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean, myXbisPstFlag As Boolean
  Private myXbisCurDirON As Boolean
    'myXbisCurDirON = False : デフォルトパスにカレントディレクトリを設定しない
    'myXbisCurDirON = True  : デフォルトパスにカレントディレクトリを設定する

'//モジュール内変数_データ
  Private myXlonFileOrgCnt As Long, myZstrFilePathOrg() As String
  Private myZvarPstData As Variant, myXobjPstFrstCell As Object
  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myXobjPstdCell As Object
    
'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    myXbisCurDirON = False
    
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
    Call callxRefSlctFilePathRptExtd
    
'//処理結果表示
    MsgBox "取得パス数：" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, myXobjDirPstdCellOUT As Object, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrDfltFldrPathIN As String, ByVal myXstrDfltFilePathIN As String, _
            ByVal myXstrExtsnIN As String, _
            ByRef myZstrAddFltrIN() As String, _
            ByVal myXbisFltrClrIN As Boolean, ByVal myXlonFltrIndxIN As Long, _
            ByVal myXlonIniViewIN As Long, ByVal myXbisMultSlctIN As Boolean, _
            ByVal myXlonOrdrCntIN As Long, ByVal myXlonTrgtWrdCntIN As Long, _
            ByRef myZvarOdrTrgtWrdPosIN() As Variant, _
            ByVal myXobjDirPstFrstCellIN As Object, ByVal myXobjFilePstFrstCellIN As Object)
    
'//入力変数を初期化
    myXlonOutputOptn = Empty
    
    myXstrDfltFldrPath = Empty: myXstrDfltFilePath = Empty
    myXstrExtsn = Empty
    Erase myZstrAddFltr: myXbisFltrClr = False: myXlonFltrIndx = Empty
    myXlonIniView = Empty: myXbisMultSlct = False
    myXlonOrdrCnt = Empty: myXlonTrgtWrdCnt = Empty: Erase myZvarOdrTrgtWrdPos
    Set myXobjDirPstFrstCell = Nothing: Set myXobjFilePstFrstCell = Nothing
    
'//入力変数を取り込み
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrDfltFldrPath = myXstrDfltFldrPathIN
    myXstrDfltFilePath = myXstrDfltFilePathIN
    myXstrExtsn = myXstrExtsnIN
    myZstrAddFltr() = myZstrAddFltrIN()
    myXbisFltrClr = myXbisFltrClrIN
    myXlonFltrIndx = myXlonFltrIndxIN
    myXlonIniView = myXlonIniViewIN
    myXbisMultSlct = myXbisMultSlctIN
    myXlonOrdrCnt = myXlonOrdrCntIN
    myXlonTrgtWrdCnt = myXlonTrgtWrdCntIN
    myZvarOdrTrgtWrdPos() = myZvarOdrTrgtWrdPosIN()
    Set myXobjDirPstFrstCell = myXobjDirPstFrstCellIN
    Set myXobjFilePstFrstCell = myXobjFilePstFrstCellIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    myXlonFileCntOUT = Empty: Erase myZobjFileOUT
    Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
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
    
'//S:ファイルパスを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
'//Run:ファイルパスをシートに書き出す
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
    myXbisCmpltFlag = False
    myXlonFileCnt = Empty: Erase myZobjFile
    Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
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
    
'    If myXlonOutputOptn < 0 And myXlonOutputOptn > 3 Then myXlonOutputOptn = 1
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables1()
        
    myXstrDfltFldrPath = "1"
    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
    
    myXstrDfltFilePath = "1"
    'myXstrDfltFilePath = ""  : デフォルトパス指定無し
    'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
    'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
    'myXstrDfltFilePath = "*" : デフォルトパスを指定
    
    myXstrExtsn = ""
    
    ReDim myZstrAddFltr(1, 2) As String
    myZstrAddFltr(1, 1) = "PDFファイル"
    myZstrAddFltr(1, 2) = "*.pdf"
    'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
    'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
    
    myXbisFltrClr = False
    'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
    'myXbisFltrClr = True  : ファイルフィルタを初期化する
    
    myXlonFltrIndx = 1
    'myXlonFltrIndx = 1〜 : ファイルフィルタの初期選択
    
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
    
    myXbisMultSlct = False
    'myXbisMultSlct = False : 複数のファイルを選択不可能
    'myXbisMultSlct = True  : 複数のファイルを選択可能
    
    myXlonOrdrCnt = 2
    myXlonTrgtWrdCnt = 2
    ReDim myZvarOdrTrgtWrdPos(myXlonOrdrCnt, myXlonTrgtWrdCnt, 2) As Variant
    myZvarOdrTrgtWrdPos(1, 1, 1) = "C"
    myZvarOdrTrgtWrdPos(1, 1, 2) = 1
    myZvarOdrTrgtWrdPos(1, 2, 1) = "Mtch"
    myZvarOdrTrgtWrdPos(1, 2, 2) = 2
    myZvarOdrTrgtWrdPos(2, 1, 1) = "C"
    myZvarOdrTrgtWrdPos(2, 1, 2) = 1
    myZvarOdrTrgtWrdPos(2, 2, 1) = "Sort"
    myZvarOdrTrgtWrdPos(2, 2, 2) = 2
    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
    
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

'SnsP_ファイルパスを取得
Private Sub snsProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt As String _
        = "ダイアログボックスが表示されますので、ファイルを選択して下さい。"
    
'//ファイルを選択してそのパスを取得
    MsgBox coXstrMsgBxPrmpt
    Call instCSlctFilePathRpt
    If myXlonFileOrgCnt <= 0 Then GoTo ExitPath
    
'//取得した2次元配列データを1次元配列データに入れ替える
  Dim myXlonTmpCnt As Long, myZstrTmp() As String
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXvarTmp As Variant, n As Long: n = L - 1
    For Each myXvarTmp In myZstrFilePathOrg
        If myXvarTmp = "" Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrTmp(n) As String
        myZstrTmp(n) = myXvarTmp
NextPath:
    Next myXvarTmp
    
'//取得したファイルパス一覧から拡張子で選別
  Dim myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String
    Call PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt, myZstrExtMtchFilePath, _
            myZstrTmp, myXstrExtsn)
    If myXlonExtMtchFileCnt <= 0 Then GoTo ExitPath
    
'//ファイルパス一覧からファイルオブジェクト一覧を取得
    Call PfixGetFileFor1DArray(myXlonFileCnt, myZobjFile, myZstrExtMtchFilePath)

'//ファイル一覧のファイル名を取得
  Dim myXlonInfoCnt As Long
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFileName, _
            myZobjFile, 1)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath

'//ファイル一覧のファイルパスを取得
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFilePath, _
            myZobjFile, 2)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExtMtchFilePath
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

'◆ClassProc名_指定文字列を含むファイル名のファイルを繰返し選択してそのパスを取得する
Private Sub instCSlctFilePathRpt()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsFilePathRpt As CSlctFilePathRpt: Set myXinsFilePathRpt = New CSlctFilePathRpt
    With myXinsFilePathRpt
    '//クラス内変数への入力
        .letDfltFldrPath = myXstrDfltFldrPath
        .letDfltFilePath = myXstrDfltFilePath
        .letExtsn = myXstrExtsn
        .letAddFltr = myZstrAddFltr
        .letFltrClr = myXbisFltrClr
        .letFltrIndx = myXlonFltrIndx
        .letIniView = myXlonIniView
        .letMultSlct = myXbisMultSlct
        .letOdrTrgtWrdPosAry = myZvarOdrTrgtWrdPos
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonFileOrgCnt = .getFileCnt
        If myXlonFileOrgCnt <= 0 Then GoTo JumpPath
        k = myXlonFileOrgCnt + Lo - 1
        ReDim myZstrFilePathOrg(k, Lo) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFileOrgCnt
            myZstrFilePathOrg(k + Lo - 1, Lo) = .getFilePathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsFilePathRpt = Nothing
    Call variablesOfCSlctFilePathRpt(myXlonFileOrgCnt, myZstrFilePathOrg)    'Debug.Print
End Sub
Private Sub variablesOfCSlctFilePathRpt( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//CSlctFilePathRptクラス内から出力した変数の内容確認
    Debug.Print "データ数: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim i As Long, j As Long
    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
        Next j
    Next i
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

 '定型Ｐ_ファイル一覧から指定拡張子と一致するファイルパスを抽出する
Private Sub PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String, _
            ByRef myXstrOrgFilePath() As String, ByVal myXstrExtsn As String)
'Includes PfncbisCheckFileExtension
'myZstrExtMtchFilePath(i) : 抽出ファイルパス
'myXstrOrgFilePath(i) : 元ファイルパス
    myXlonExtMtchFileCnt = Empty: Erase myZstrExtMtchFilePath
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myXstrOrgFilePath): myXstrTmp = myXstrOrgFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarFilePath As Variant, myXbisExtChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarFilePath In myXstrOrgFilePath
      Dim myXstrFilePath As String: myXstrFilePath = myXvarFilePath
        myXbisExtChck = PfncbisCheckFileExtension(myXstrFilePath, myXstrExtsn)
        If myXbisExtChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExtMtchFilePath(n) As String
        myZstrExtMtchFilePath(n) = myXvarFilePath
NextPath:
    Next
    myXlonExtMtchFileCnt = n - Lo + 1
    myXvarFilePath = Empty
ExitPath:
End Sub

 '定型Ｆ_指定ファイルが指定拡張子であることを確認する
Private Function PfncbisCheckFileExtension( _
            ByVal myXstrFilePath As String, ByVal myXstrExtsn As String) As Boolean
'myXstrExtsn = "*" : 任意の文字列のワイルドカード
    PfncbisCheckFileExtension = False
    If myXstrFilePath = "" Then Exit Function
    If myXstrExtsn = "" Then GoTo JumpPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrOrgExt As String
    With myXobjFSO
        If .FileExists(myXstrFilePath) = False Then Exit Function
        myXstrOrgExt = .GetExtensionName(myXstrFilePath)
    End With
  Dim myXstrDesExt As String: myXstrDesExt = myXstrExtsn
    If Left(myXstrDesExt, 1) = "." Then myXstrDesExt = Mid(myXstrDesExt, 2)
    myXstrOrgExt = LCase(myXstrOrgExt)
    myXstrDesExt = LCase(myXstrDesExt)
    If myXstrOrgExt = myXstrDesExt Then GoTo JumpPath
  Dim myXlonPstn As Long: myXlonPstn = InStr(myXstrDesExt, "*")
    Select Case myXlonPstn
        Case 1
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
        Case Len(myXstrExtsn)
            If Left(myXstrOrgExt, Len(myXstrDesExt) - 1) _
                    <> Left(myXstrDesExt, Len(myXstrDesExt) - 1) Then _
                Exit Function
        Case Else
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
            If Left(myXstrOrgExt, myXlonPstn - 1) _
                    <> Left(myXstrDesExt, myXlonPstn - 1) Then _
                Exit Function
    End Select
    Set myXobjFSO = Nothing
JumpPath:
    PfncbisCheckFileExtension = True
End Function

 '定型Ｐ_1次元配列のファイルパスからファイルオブジェクト一覧を取得する
Private Sub PfixGetFileFor1DArray( _
                myXlonFileCnt As Long, myZobjFile() As Object, _
                ByRef myZstrFilePath() As String)
'myZobjFile(i) : ファイルオブジェクト一覧
'myZstrFilePath(i) : 元ファイルパス一覧
    myXlonFileCnt = Empty: Erase myZobjFile
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFilePath): myXstrTmp = myZstrFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjTmp As Object, i As Long, n As Long: n = Lo - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath) To UBound(myZstrFilePath)
        myXstrTmp = Empty
        myXstrTmp = myZstrFilePath(i)
        With myXobjFSO
            If .FileExists(myXstrTmp) = False Then GoTo NextPath
            Set myXobjTmp = .GetFile(myXstrTmp)
        End With
        n = n + 1: ReDim Preserve myZobjFile(n) As Object
        Set myZobjFile(n) = myXobjTmp
NextPath:
    Next i
    myXlonFileCnt = n - Lo + 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '定型Ｐ_1次元配列のフォルダファイルオブジェクト一覧の文字列情報を取得する
Private Sub PfixGetFolderFileStringInformationFor1DArray( _
                myXlonInfoCnt As Long, myZstrInfo() As String, _
                ByRef myZobjFldrFile() As Object, _
                Optional ByVal coXlonStrOptn As Long = 1)
'myZstrInfo(i) : 抽出フォルダ情報
'myZobjFldrFile(i) : 元フォルダor元ファイル
'coXlonStrOptn = 1  : 名前 (Name)
'coXlonStrOptn = 2  : パス (Path)
'coXlonStrOptn = 3  : 親フォルダ (ParentFolder)
'coXlonStrOptn = 4  : 属性 (Attributes)
'coXlonStrOptn = 5  : 種類 (Type)
    myXlonInfoCnt = Empty: Erase myZstrInfo
  Dim myXobjTmp As Object, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXstrTmp As String, i As Long, n As Long: n = Lo - 1
    On Error GoTo NextPath
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        myXstrTmp = Empty
        If myZobjFldrFile(i) Is Nothing Then GoTo NextPath
        Select Case coXlonStrOptn
            Case 1: myXstrTmp = myZobjFldrFile(i).Name
            Case 2: myXstrTmp = myZobjFldrFile(i).Path
            Case 3: myXstrTmp = myZobjFldrFile(i).ParentFolder
            Case 4: myXstrTmp = myZobjFldrFile(i).Attributes
            Case 5: myXstrTmp = myZobjFldrFile(i).Type
        End Select
        n = n + 1: ReDim Preserve myZstrInfo(n) As String
        myZstrInfo(n) = myXstrTmp
NextPath:
    Next i
    On Error GoTo 0
    myXlonInfoCnt = n - Lo + 1
ExitPath:
End Sub

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
'Private Sub setControlVariables()
'    myXstrDfltFldrPath = "1"
'    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
'    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
'    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
'    myXstrDfltFilePath = "1"
'    'myXstrDfltFilePath = ""  : デフォルトパス指定無し
'    'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "*" : デフォルトパスを指定
'    myXstrExtsn = ""
'    ReDim myZstrAddFltr(1, 2) As String
'    myZstrAddFltr(1, 1) = "PDFファイル"
'    myZstrAddFltr(1, 2) = "*.pdf"
'    'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
'    'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
'    myXbisFltrClr = False
'    'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
'    'myXbisFltrClr = True  : ファイルフィルタを初期化する
'    myXlonFltrIndx = 1
'    'myXlonFltrIndx = 1〜 : ファイルフィルタの初期選択
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
'    myXbisMultSlct = False
'    'myXbisMultSlct = False : 複数のファイルを選択不可能
'    'myXbisMultSlct = True  : 複数のファイルを選択可能
'    myXlonOrdrCnt = 2
'    myXlonTrgtWrdCnt = 2
'    ReDim myZvarOdrTrgtWrdPos(myXlonOrdrCnt, myXlonTrgtWrdCnt, 2) As Variant
'    myZvarOdrTrgtWrdPos(1, 1, 1) = "C"
'    myZvarOdrTrgtWrdPos(1, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(1, 2, 1) = "Mtch"
'    myZvarOdrTrgtWrdPos(1, 2, 2) = 2
'    myZvarOdrTrgtWrdPos(2, 1, 1) = "C"
'    myZvarOdrTrgtWrdPos(2, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(2, 2, 1) = "Sort"
'    myZvarOdrTrgtWrdPos(2, 2, 2) = 2
'    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
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
'    myXbisInptBxOFF = False
'    'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
'    'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
'    myXbisEachWrtON = False
'    'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
'    'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
'End Sub
'◆ModuleProc名_指定文字列を含むファイル名のファイルを繰返し選択してそのパスを取得してシートに書き出す
Private Sub callxRefSlctFilePathRptExtd()
'  Dim myXlonOutputOptn As Long, _
'        myXstrDfltFldrPath As String, myXstrDfltFilePath As String, myXstrExtsn As String, _
'        myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long, _
'        myXlonIniView As Long, myXbisMultSlct As Boolean, _
'        myXlonOrdrCnt As Long, myXlonTrgtWrdCnt As Long, myZvarOdrTrgtWrdPos() As Variant, _
'        myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : ファイルパスをエクセルシートに書き出す
'    'myXlonOutputOptn = 2 : ファイル名をエクセルシートに書き出す
'    'myXlonOutputOptn = 3 : 親フォルダに応じてファイルパス／名をエクセルシートに書き出す
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
'    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す
'    'myXlonOutputOptn = 3 : 親フォルダに応じてフォルダパス／名をエクセルシートに書き出す
'    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
'    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
'    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
'    'myXstrDfltFilePath = ""  : デフォルトパス指定無し
'    'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "*" : デフォルトパスを指定
'    'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
'    'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
'    'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
'    'myXbisFltrClr = True  : ファイルフィルタを初期化する
'    'myXlonFltrIndx = 1〜 : ファイルフィルタの初期選択
'    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
'    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
'    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
'    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
'    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
'    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
'    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
'    'myXbisMultSlct = False : 複数のファイルを選択不可能
'    'myXbisMultSlct = True  : 複数のファイルを選択可能
'    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
'    'myZobjFile(i) : ファイルオブジェクト
'    'myZstrFileName(i) : ファイル名
'    'myZstrFilePath(i) : ファイルパス
    Call xRefSlctFilePathRptExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXlonOutputOptn, _
            myXstrDfltFldrPath, myXstrDfltFilePath, myXstrExtsn, _
            myZstrAddFltr, myXbisFltrClr, myXlonFltrIndx, _
            myXlonIniView, myXbisMultSlct, _
            myXlonOrdrCnt, myXlonTrgtWrdCnt, myZvarOdrTrgtWrdPos, _
            myXobjDirPstFrstCell, myXobjFilePstFrstCell)
    Debug.Print "データ: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSlctFilePathRptExtd()
'//xRefSlctFilePathRptExtdモジュールのモジュールメモリのリセット処理
    Call xRefSlctFilePathRptExtd.resetConstant
End Sub
