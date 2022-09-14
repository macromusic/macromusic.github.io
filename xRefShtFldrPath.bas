Attribute VB_Name = "xRefShtFldrPath"
'Includes CSrchShtCmnt
'Includes PfncbisCheckFolderExist
'Includes PfncobjGetFolder
'Includes PfixGetFolderNameInformationByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルシート上に記載されたフォルダパスを取得する
'Rev.003
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefShtFldrPath"
  Private Const meMlonExeNum As Long = 0
  
'//出力データ
  Private myXobjFldr As Object, myXstrFldrName As String, myXstrFldrPath As String, _
            myXstrDirPath As String, _
            myXobjFldrPstdFrstCell As Object
  
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
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
'    'myZstrTrgtVal(i) : 取得文字列
'    'myZobjTrgtRng(i) : 行列位置のセル

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonTrgtValCnt = Empty: Erase myZstrTrgtVal: Erase myZobjTrgtRng
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
    Call callOfxRefShtFldrPath
    
'//処理結果表示
    MsgBox "取得パス：" & myXstrFldrPath
    
End Sub

'PublicP_
Public Sub callProc( _
            myXobjFldrOUT As Object, myXstrFldrNameOUT As String, myXstrFldrPathOUT As String, _
            myXstrDirPathOUT As String, _
            myXobjFldrPstdFrstCellOUT As Object, _
            ByVal myXlonSrchShtNoIN As Long, ByVal myXobjSrchSheetIN As Object, _
            ByVal myXlonShtSrchCntIN As Long, ByRef myZvarSrchCndtnIN As Variant, _
            ByVal myXbisInStrOptnIN As Boolean)
    
'//入力変数を初期化
    myXlonSrchShtNo = Empty: Set myXobjSrchSheet = Nothing
    myXlonShtSrchCnt = Empty: myZvarSrchCndtn = Empty
    myXbisInStrOptn = False

'//入力変数を取り込み
    myXlonSrchShtNo = myXlonSrchShtNoIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    myXlonShtSrchCnt = myXlonShtSrchCntIN
    myZvarSrchCndtn = myZvarSrchCndtnIN
    myXbisInStrOptn = myXbisInStrOptnIN
    
'//出力変数を初期化
    Set myXobjFldrOUT = Nothing: myXstrFldrNameOUT = Empty: myXstrFldrPathOUT = Empty
    myXstrDirPathOUT = Empty
    Set myXobjFldrPstdFrstCellOUT = Nothing
    
'//処理実行
    Call ctrProc
    If myXstrFldrPath = "" Then Exit Sub
    
'//出力変数に格納
    Set myXobjFldrOUT = myXobjFldr
    myXstrFldrNameOUT = myXstrFldrName
    myXstrFldrPathOUT = myXstrFldrPath
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjFldrPstdFrstCellOUT = myXobjFldrPstdFrstCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:シート上の記載データを取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXstrFldrPath = Empty: myXstrFldrName = Empty
    myXstrDirPath = Empty: Set myXobjFldr = Nothing
    Set myXobjFldrPstdFrstCell = Nothing
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
    
    myXlonSrchShtNo = 4
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'    Set myXobjSrchSheet = ActiveSheet

  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 1
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : 検索文字列
    'myZvarSrchCndtn(i, 2) : オフセット行数
    'myZvarSrchCndtn(i, 3) : オフセット列数
    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "フォルダパス："
    myZvarSrchCndtn(k, L + 1) = 0
    myZvarSrchCndtn(k, L + 2) = 1
    myZvarSrchCndtn(k, L + 3) = 0
    
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
    
End Sub

'SnsP_シート上の記載データを取得
Private Sub snsProc()
    myXbisExitFlag = False
    
'//フォルダパスを検索して取得
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
  Dim L As Long: L = LBound(myZobjTrgtRng)
    myXstrFldrPath = myZstrTrgtVal(L)
    Set myXobjFldrPstdFrstCell = myZobjTrgtRng(L)
    If myXobjFldrPstdFrstCell Is Nothing Then GoTo ExitPath
    
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

'PrcsP_取得データ内容をチェック
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
'    myXlonSrchShtNo = 3
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
''    Set myXobjSrchSheet = ActiveSheet
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'    myXlonShtSrchCnt = 1
'    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
'  Dim k As Long: k = L - 1
'    k = k + 1   'k = 1
'    myZvarSrchCndtn(k, L + 0) = "フォルダパス："
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
'End Sub
'◆ModuleProc名_エクセルシート上に記載されたフォルダパスを取得する
Private Sub callOfxRefShtFldrPath()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'    'myZvarSrchCndtn(i, 1) : 検索文字列
'    'myZvarSrchCndtn(i, 2) : オフセット行数
'    'myZvarSrchCndtn(i, 3) : オフセット列数
'    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
'    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
'    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
'  Dim myXstrFldrPath As String, myXstrFldrName As String, _
'        myXstrDirPath As String, myXobjFldr As Object, _
'        myXobjFldrPstdFrstCell As Object
    Call xRefShtFldrPath.callProc( _
            myXobjFldr, myXstrFldrName, myXstrFldrPath, myXstrDirPath, _
            myXobjFldrPstdFrstCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn)
    Debug.Print "データ: " & myXstrFldrPath
    Debug.Print "データ: " & myXstrFldrName
    Debug.Print "データ: " & myXstrDirPath
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefShtFldrPath()
'//xRefShtFldrPathモジュールのモジュールメモリのリセット処理
    Call xRefShtFldrPath.resetConstant
End Sub
