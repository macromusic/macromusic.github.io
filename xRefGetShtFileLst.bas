Attribute VB_Name = "xRefGetShtFileLst"
'Includes CSlctShtSrsData
'Includes CSlctShtDscrtData
'Includes PfixPickUpExistFilePathArray
'Includes PfixGetFileNameArrayByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルシート上に記載されたデータを選択してパス一覧を取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefGetShtFileLst"
  Private Const meMlonExeNum As Long = 0
  
'//出力データ
  Private myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
    'myZstrFileName(k) : ファイル名
    'myZstrFilePath(k) : ファイルパス
  
'//入力データ
  Private myXbisByDscrt As Boolean
  Private myXlonRngOptn As Long
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
    'myXbisByDscrt = False : シート上の連続範囲を指定して取得する
    'myXbisByDscrt = True  : シート上の不連続範囲を指定して取得する
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonDataRowCnt As Long, myXlonDataColCnt As Long, myXlonDataCnt As Long, _
            myZstrShtData() As String, myZvarShtData() As Variant
  Private myZstrFilePathOrg() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonDataRowCnt = Empty: myXlonDataColCnt = Empty: myXlonDataCnt = Empty
    Erase myZstrShtData: Erase myZvarShtData
    Erase myZstrFilePathOrg
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
    Call callxRefGetShtFileLst
    
'//処理結果表示
    MsgBox "取得パス数：" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFileCntOUT As Long, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String)
    
'//出力変数を初期化
    myXlonFileCntOUT = Empty: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    
'//処理実行
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//出力変数に格納
    myXlonFileCntOUT = myXlonFileCnt
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:シート上の記載データを取得
    Select Case myXbisByDscrt
        Case True: Call snsProc2
        Case Else: Call snsProc1
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:取得データ内容をチェック
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty: Erase myZstrFileName: Erase myZstrFilePath
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
    
    myXbisByDscrt = True
    'myXbisByDscrt = False : シート上の連続範囲を指定して取得する
    'myXbisByDscrt = True  : シート上の不連続範囲を指定して取得する
    
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
    
    myXstrInptBxPrmpt = "処理したいファイルパスを選択して下さい。"
    myXstrInptBxTtl = "ファイルパスの選択"
    
End Sub

'SnsP_シート上の記載データを取得
Private Sub snsProc1()
    myXbisExitFlag = False
    
'//シート上の連続範囲を指定してその範囲のデータと情報を取得
    Call instCSlctShtSrsData
    If myXlonDataRowCnt <= 0 Or myXlonDataColCnt <= 0 Then GoTo ExitPath
    
  Dim i As Long, j As Long, k As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonDataCnt = myXlonDataRowCnt * myXlonDataColCnt
    k = myXlonDataCnt + L - 1
    ReDim myZstrFilePathOrg(k) As String
    k = L - 1
    For j = LBound(myZstrShtData, 2) To UBound(myZstrShtData, 2)
        For i = LBound(myZstrShtData, 1) To UBound(myZstrShtData, 1)
            k = k + 1
            myZstrFilePathOrg(k) = myZstrShtData(i, j)
        Next i
    Next j
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_シート上の記載データを取得
Private Sub snsProc2()
    myXbisExitFlag = False
  
'//シート上の不連続範囲を指定してその範囲のデータと情報を取得
    Call instCSlctShtDscrtData
    If myXlonDataCnt <= 0 Then GoTo ExitPath
    
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim k As Long
    k = UBound(myZvarShtData, 1)
    ReDim myZstrFilePathOrg(k) As String
    For k = LBound(myZvarShtData, 1) To UBound(myZvarShtData, 1)
        myZstrFilePathOrg(k) = CStr(myZvarShtData(k, L + 2))
    Next k
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_取得データ内容をチェック
Private Sub prsProc()
    myXbisExitFlag = False
    
'//ファイルパス一覧から存在するファイルパスを抽出
    Call PfixPickUpExistFilePathArray(myXlonFileCnt, myZstrFilePath, myZstrFilePathOrg)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
'//指定ファイルパス一覧のファイル名一覧を取得
    Call PfixGetFileNameArrayByFSO(myXlonFileCnt, myZstrFileName, myZstrFilePath)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
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

'◆ClassProc名_シート上の連続範囲を指定してその範囲のデータと情報を取得する
Private Sub instCSlctShtSrsData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSlctShtSrsData As CSlctShtSrsData: Set myXinsSlctShtSrsData = New CSlctShtSrsData
    With myXinsSlctShtSrsData
    '//クラス内変数への入力
        .letRngOptn = myXlonRngOptn
        .letByVrnt = False
        .letGetCmnt = False
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonDataRowCnt = .getDataRowCnt
        myXlonDataColCnt = .getDataColCnt
        If myXlonDataRowCnt <= 0 Or myXlonDataColCnt <= 0 Then GoTo JumpPath
        i = myXlonDataRowCnt + Lo - 1: j = myXlonDataColCnt + Lo - 1
        ReDim myZstrShtData(i, j) As String
        Lc = .getOptnBase
        For j = 1 To myXlonDataColCnt
            For i = 1 To myXlonDataRowCnt
                myZstrShtData(i + Lo - 1, j + Lo - 1) _
                    = .getStrShtDataAry(i + Lc - 1, j + Lc - 1)
            Next i
        Next j
    End With
JumpPath:
    Set myXinsSlctShtSrsData = Nothing
End Sub

'◆ClassProc名_シート上の不連続範囲を指定してその範囲のデータと情報を取得する
Private Sub instCSlctShtDscrtData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsSlctShtDscrtData As CSlctShtDscrtData
    Set myXinsSlctShtDscrtData = New CSlctShtDscrtData
    With myXinsSlctShtDscrtData
    '//クラス内変数への入力
        .letByVrnt = False
        .letGetCmnt = False
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonDataCnt = .getDataCnt
        If myXlonDataCnt <= 0 Then GoTo JumpPath
        i = myXlonDataCnt + Lo - 1
        ReDim myZvarShtData(i, Lo + 3) As Variant
        Lc = .getOptnBase
        For i = 1 To myXlonDataCnt
            myZvarShtData(i + Lo - 1, Lo + 2) = .getShtDataAry(i + Lc - 1, Lc + 2)
        Next i
    End With
JumpPath:
    Set myXinsSlctShtDscrtData = Nothing
End Sub

'===============================================================================================

 '定型Ｐ_ファイルパス一覧から存在するファイルパスを抽出する
Private Sub PfixPickUpExistFilePathArray( _
            myXlonExistFileCnt As Long, myZstrExistFilePath() As String, _
            ByRef myZstrOrgFilePath() As String)
'myZstrExistFilePath(i) : 抽出ファイルパス
'myZstrOrgFilePath(i) : 元ファイルパス
    myXlonExistFileCnt = Empty: Erase myZstrExistFilePath
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrOrgFilePath): myXstrTmp = myZstrOrgFilePath(Li)
    On Error GoTo 0
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarPath As Variant, myXbisExistChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarPath In myZstrOrgFilePath
        myXbisExistChck = myXobjFSO.FileExists(myXvarPath)
        If myXbisExistChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExistFilePath(n) As String
        myZstrExistFilePath(n) = CStr(myXvarPath)
NextPath:
    Next myXvarPath
    myXlonExistFileCnt = n - Lo + 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '定型Ｐ_指定ファイルパス一覧のファイル名一覧を取得する(FileSystemObject使用)
Private Sub PfixGetFileNameArrayByFSO( _
            myXlonFileCnt As Long, myZstrFileName() As String, _
            ByRef myZstrFilePath() As String)
'myZstrFileName(i) : ファイル名
'myZstrFilePath(i) : ファイルパス
    myXlonFileCnt = Empty: Erase myZstrFileName
  Dim myXstrTmp As String, Li As Long, Ui As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFilePath): Ui = UBound(myZstrFilePath)
    myXstrTmp = myZstrFilePath(Li)
    On Error GoTo 0
    myXlonFileCnt = Ui - Li + 1
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, myXbisFileExist As Boolean
    i = myXlonFileCnt + Lo - 1: ReDim myZstrFileName(i) As String
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath) To UBound(myZstrFilePath)
        myXbisFileExist = myXobjFSO.FileExists(myZstrFilePath(i))
        If myXbisFileExist = True Then _
            myZstrFileName(i) = myXobjFSO.getFileName(myZstrFilePath(i))
    Next i
    Set myXobjFSO = Nothing
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
'Private Sub setControlVariables()
'    myXbisByDscrt = False
'    'myXbisByDscrt = False : シート上の連続範囲を指定して取得する
'    'myXbisByDscrt = True  : シート上の不連続範囲を指定して取得する
'    myXlonRngOptn = 0
'    myXlonRngOptn = 0
'    'myXlonRngOptn = 0  : 選択範囲
'    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
'    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
'    'myXlonRngOptn = 3  : 全データ範囲
'    myXstrInptBxPrmpt = "処理したいファイルパスを選択して下さい。"
'    myXstrInptBxTtl = "ファイルパスの選択"
'End Sub
'◆ModuleProc名_エクセルシート上に記載されたデータを選択してパス一覧を取得する
Private Sub callxRefGetShtFileLst()
'  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
'    'myZstrFileName(k) : ファイル名
'    'myZstrFilePath(k) : ファイルパス
    Call xRefGetShtFileLst.callProc( _
            myXlonFileCnt, myZstrFileName, myZstrFilePath)
    Call variablesOfxRefGetShtFileLst(myXlonFileCnt, myZstrFilePath) 'Debug.Print
End Sub
Private Sub variablesOfxRefGetShtFileLst( _
            myXlonDataCnt As Long, myXvarField As Variant)
'//xRefGetShtFileLst内から出力した変数の内容確認
    Debug.Print "データ数: " & myXlonDataCnt
    If myXlonDataCnt = 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myXvarField) To UBound(myXvarField)
        Debug.Print "データ" & k & ": " & myXvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefGetShtFileLst()
'//xRefGetShtFileLstモジュールのモジュールメモリのリセット処理
    Call xRefGetShtFileLst.resetConstant
End Sub
