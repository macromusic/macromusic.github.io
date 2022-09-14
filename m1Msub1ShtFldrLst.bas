Attribute VB_Name = "m1Msub1ShtFldrLst"
'Includes CSrchShtCmnt
'Includes CSeriesData
'Includes PfixPickUpExistFolderArray
'Includes PfixGetFolderFor1DArray
'Includes PfixGetFolderFileStringInformationFor1DArray
'Includes PfixChangeModuleConstValue

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルシート上に記載されたフォルダパス一覧を取得する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "m1Msub1ShtFldrLst"
  Private Const meMlonExeNum As Long = 0
  
'//出力データ
  Private myXlonFldrCnt As Long, myZobjFldr() As Object, _
            myZstrFldrName() As String, myZstrFldrPath() As String, _
            myXobjFldrPstdFrstCell As Object
    'myZobjFldr(k) : フォルダオブジェクト
    'myZstrFldrName(k) : フォルダ名
    'myZstrFldrPath(k) : フォルダパス
  Private myXstrDirPath As String, myXobjDirPstdCell As Object
  
'//入力データ
  Private myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
            myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
            myXbisInStrOptn As Boolean
    'myZvarSrchCndtn(i, 1) : 検索文字列
    'myZvarSrchCndtn(i, 2) : オフセット行数
    'myZvarSrchCndtn(i, 3) : オフセット列数
    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
  
  Private myXbisRowDrctn As Boolean
    'myXbisRowDrctn = True  : 行方向のみを検索
    'myXbisRowDrctn = False : 列方向のみを検索
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
'    'myZstrTrgtVal(i) : 取得文字列
'    'myZobjTrgtRng(i) : 行列位置のセル
  
  Private myXlonBgnRow As Long, myXlonBgnCol As Long
  
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : 取得文字列

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonTrgtValCnt = Empty: Erase myZstrTrgtVal: Erase myZobjTrgtRng
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
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
    Call callm1Msub1ShtFldrLst
    
'//処理結果表示
    MsgBox "取得パス数：" & myXlonFldrCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            myXobjFldrPstdFrstCellOUT As Object, _
            myXstrDirPathOUT As String, myXobjDirPstdCellOUT As Object, _
            ByVal myXlonSrchShtNoIN As Long, ByVal myXobjSrchSheetIN As Object, _
            ByVal myXlonShtSrchCntIN As Long, ByRef myZvarSrchCndtnIN As Variant, _
            ByVal myXbisInStrOptnIN As Boolean, _
            ByVal myXbisRowDrctnIN As Boolean)
    
'//入力変数を初期化
    myXlonSrchShtNo = Empty: Set myXobjSrchSheet = Nothing
    myXlonShtSrchCnt = Empty: myZvarSrchCndtn = Empty
    myXbisInStrOptn = False
    myXbisRowDrctn = False

'//入力変数を取り込み
    myXlonSrchShtNo = myXlonSrchShtNoIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    myXlonShtSrchCnt = myXlonShtSrchCntIN
    myZvarSrchCndtn = myZvarSrchCndtnIN
    myXbisInStrOptn = myXbisInStrOptnIN
    myXbisRowDrctn = myXbisRowDrctnIN
    
'//出力変数を初期化
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT: Erase myZstrFldrNameOUT: Erase myZobjFldrOUT
    Set myXobjFldrPstdFrstCellOUT = Nothing
    myXstrDirPathOUT = Empty: Set myXobjDirPstdCellOUT = Nothing
    
'//処理実行
    Call ctrProc
'    If myXlonFldrCnt <= 0 Then Exit Sub
    
'//出力変数に格納
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    Set myXobjFldrPstdFrstCellOUT = myXobjFldrPstdFrstCell
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables1
    Call setControlVariables2
    
'//S:シート上の記載データを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFldrCnt = Empty: Erase myZobjFldr: Erase myZstrFldrName: Erase myZstrFldrPath
    Set myXobjFldrPstdFrstCell = Nothing
    myXstrDirPath = Empty: Set myXobjDirPstdCell = Nothing
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
Private Sub setControlVariables1()
    
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    Set myXobjSrchSheet = ActiveSheet

  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 2
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : 検索文字列
    'myZvarSrchCndtn(i, 2) : オフセット行数
    'myZvarSrchCndtn(i, 3) : オフセット列数
    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "親フォルダパス："
    myZvarSrchCndtn(k, L + 1) = 0
    myZvarSrchCndtn(k, L + 2) = 1
    myZvarSrchCndtn(k, L + 3) = 0
    k = k + 1   'k = 2
    myZvarSrchCndtn(k, L + 0) = "フォルダ一覧"
    myZvarSrchCndtn(k, L + 1) = 1
    myZvarSrchCndtn(k, L + 2) = 0
    myZvarSrchCndtn(k, L + 3) = 0
    
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : 行方向のみを検索
    'myXbisRowDrctn = False : 列方向のみを検索
    
End Sub

'SnsP_シート上の記載データを取得
Private Sub snsProc()
    myXbisExitFlag = False
    
'//ディレクトリパスを検索して取得
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim Lc As Long: Lc = LBound(myZstrTrgtVal)
    myXstrDirPath = myZstrTrgtVal(Lc + 0)
    Set myXobjDirPstdCell = myZobjTrgtRng(Lc + 0)
    Set myXobjFldrPstdFrstCell = myZobjTrgtRng(Lc + 1)
    If myXobjFldrPstdFrstCell Is Nothing Then GoTo ExitPath
    
'//フォルダパス一覧を取得
    myXlonBgnRow = myXobjFldrPstdFrstCell.Row
    myXlonBgnCol = myXobjFldrPstdFrstCell.Column
    Call instCSeriesData
    If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
    
  Dim myZstrFldrPathOrg1() As String, myZstrFldrPathOrg2() As String
  Dim i As Long
    i = myXlonSrsDataCnt + Lo - 1
    ReDim myZstrFldrPathOrg1(i) As String
    ReDim myZstrFldrPathOrg2(i) As String
    Lc = LBound(myZstrSrsData)
    For i = 1 To myXlonSrsDataCnt
        myZstrFldrPathOrg1(i + Lo - 1) = myXstrDirPath & "\" & myZstrSrsData(i + Lc - 1)
        myZstrFldrPathOrg2(i + Lo - 1) = myZstrSrsData(i + Lc - 1)
    Next i
    
'//取得したフォルダパス一覧から存在で選別
  Dim myXlonExistFldrCnt As Long, myZstrExistFldrPath() As String
    Call PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt, myZstrExistFldrPath, _
            myZstrFldrPathOrg1)
    If myXlonExistFldrCnt > 0 Then GoTo JumpPath
    
    Call PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt, myZstrExistFldrPath, _
            myZstrFldrPathOrg2)
    If myXlonExistFldrCnt <= 0 Then GoTo ExitPath
    
JumpPath:
'//フォルダパス一覧からフォルダオブジェクト一覧を取得
    Call PfixGetFolderFor1DArray(myXlonFldrCnt, myZobjFldr, myZstrExistFldrPath)

'//フォルダ一覧のフォルダ名を取得
  Dim myXlonInfoCnt As Long
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFldrName, _
            myZobjFldr, 1)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath

'//フォルダ一覧のフォルダパスを取得
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFldrPath, _
            myZobjFldr, 2)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrTrgtVal: Erase myZobjTrgtRng
    Erase myZstrFldrPathOrg1: Erase myZstrFldrPathOrg2
    Erase myZstrExistFldrPath
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

'◆ClassProc名_シート上のデータから文字列を検索してデータと位置情報を取得する
Private Sub instCSrchShtCmnt()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSrchShtCmnt As CSrchShtCmnt: Set myXinsSrchShtCmnt = New CSrchShtCmnt
    With myXinsSrchShtCmnt
    '//文字列検索シートと検索条件を設定
        Set .setSrchSheet = myXobjSrchSheet
        .letSrchCndtn = myZvarSrchCndtn
        .letInStrOptn = myXbisInStrOptn
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonTrgtValCnt = .getValCnt
        If myXlonTrgtValCnt <= 0 Then GoTo JumpPath
        i = myXlonTrgtValCnt + Lo - 1: j = Lo + 1
        ReDim myZstrTrgtVal(i) As String
        ReDim myZobjTrgtRng(i) As Object
        Lc = .getOptnBase
        For i = 1 To myXlonTrgtValCnt
            myZstrTrgtVal(i + Lo - 1) = .getValAry(i + Lc - 1)
            Set myZobjTrgtRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
        Next i
    End With
JumpPath:
    Set myXinsSrchShtCmnt = Nothing
End Sub

'◆ClassProc名_シート上の連続するデータ範囲を取得する
Private Sub instCSeriesData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
    With myXinsSeriesData
    '//クラス内変数への入力
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
        .letRowDrctn = myXbisRowDrctn
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonSrsDataCnt = .getSrsDataCnt
        If myXlonSrsDataCnt <= 0 Then GoTo JumpPath
        k = myXlonSrsDataCnt + Lo - 1
        ReDim myZstrSrsData(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonSrsDataCnt
            myZstrSrsData(k + Lo - 1) = .getSrsDataAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSeriesData = Nothing
End Sub

'===============================================================================================

 '定型Ｐ_フォルダパス一覧から存在するフォルダパスを抽出する
Private Sub PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt As Long, myZstrExistFldrPath() As String, _
            ByRef myZstrOrgFldrPath() As String)
'myZstrExistFldrPath(i) : 抽出フォルダパス
'myZstrOrgFldrPath(i) : 元フォルダパス
    myXlonExistFldrCnt = Empty: Erase myZstrExistFldrPath
  Dim myXstrTmp As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgFldrPath): myXstrTmp = myZstrOrgFldrPath(L)
    On Error GoTo 0
  Dim i As Long, myXstrPath As String, n As Long: n = L - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrOrgFldrPath) To UBound(myZstrOrgFldrPath)
        myXstrPath = myZstrOrgFldrPath(i)
        If myXobjFSO.FolderExists(myXstrPath) = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExistFldrPath(n) As String
        myZstrExistFldrPath(n) = myXstrPath
NextPath:
    Next i
    myXlonExistFldrCnt = n + L - 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '定型Ｐ_1次元配列のフォルダパス一覧からフォルダオブジェクト一覧を取得する
Private Sub PfixGetFolderFor1DArray( _
                myXlonFldrCnt As Long, myZobjFldr() As Object, _
                ByRef myZstrFldrPath() As String)
'myZobjFldr(i) : フォルダオブジェクト一覧
'myZstrFldrPath(i) : 元フォルダパス一覧
    myXlonFldrCnt = Empty: Erase myZobjFldr
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFldrPath): myXstrTmp = myZstrFldrPath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjTmp As Object, i As Long, n As Long: n = Lo - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFldrPath) To UBound(myZstrFldrPath)
        myXstrTmp = Empty
        myXstrTmp = myZstrFldrPath(i)
        With myXobjFSO
            If .FolderExists(myXstrTmp) = False Then GoTo NextPath
            Set myXobjTmp = .GetFolder(myXstrTmp)
        End With
        n = n + 1: ReDim Preserve myZobjFldr(n) As Object
        Set myZobjFldr(n) = myXobjTmp
NextPath:
    Next i
    myXlonFldrCnt = n - Lo + 1
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
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
''    Set myXobjSrchSheet = ActiveSheet
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'    myXlonShtSrchCnt = 2
'    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
'  Dim k As Long: k = L - 1
'    k = k + 1   'k = 1
'    myZvarSrchCndtn(k, L + 0) = "親フォルダパス："
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    k = k + 1   'k = 2
'    myZvarSrchCndtn(k, L + 0) = "フォルダ一覧"
'    myZvarSrchCndtn(k, L + 1) = 1
'    myZvarSrchCndtn(k, L + 2) = 0
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    myXbisRowDrctn = True
'    'myXbisRowDrctn = True  : 行方向のみを検索
'    'myXbisRowDrctn = False : 列方向のみを検索
'End Sub
'◆ModuleProc名_エクセルシート上に記載されたフォルダパス一覧を取得する
Private Sub callm1Msub1ShtFldrLst()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
'  Dim myXbisRowDrctn As Boolean
'    'myXbisRowDrctn = True  : 行方向のみを検索
'    'myXbisRowDrctn = False : 列方向のみを検索
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String, _
'        myXobjFldrPstdFrstCell As Object, _
'        myXstrDirPath As String, myXobjDirPstdCell As Object
'    'myZobjFldr(k) : フォルダオブジェクト
'    'myZstrFldrName(k) : フォルダ名
'    'myZstrFldrPath(k) : フォルダパス
    Call m1Msub1ShtFldrLst.callProc( _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, _
            myXobjFldrPstdFrstCell, _
            myXstrDirPath, myXobjDirPstdCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn, myXbisRowDrctn)
'    Debug.Print "データ: " & myXlonFldrCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInm1Msub1ShtFldrLst()
'//m1Msub1ShtFldrLstモジュールのモジュールメモリのリセット処理
    Call m1Msub1ShtFldrLst.resetConstant
End Sub
