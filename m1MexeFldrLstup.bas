Attribute VB_Name = "m1MexeFldrLstup"
'Includes m1Msub1ShtFldrLst
'Includes m1Msub2SlctFldrPathExtd
'Includes m1Msub3SubFldrLstExtd
'Includes PfncbisCheckFolderExist
'Includes PfixChangeModuleConstValue

Option Explicit
Option Base 1

'◆ModuleProc名_処理フォルダをリストアップする
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "m1MexeFldrLstup"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
        myZstrFldrName() As String, myZstrFldrPath() As String, _
        myXobjFldrPstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object
    'myZobjFldr(k) : フォルダオブジェクト
    'myZstrFldrName(k) : フォルダ名
    'myZstrFldrPath(k) : フォルダパス
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Dim myXbisDirPstFlag As Boolean
    'myXbisDirPstFlag = True  : 親フォルダパスの貼り付け有り
    'myXbisDirPstFlag = False : 親フォルダパスの貼り付け無し
  
'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisDirPstFlag = False
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
    '処理:  '◆ModuleProc名_エクセルシート上に記載されたフォルダパス一覧を取得する
            '◆ModuleProc名_フォルダを選択してそのパスを取得してシートに書き出す
            '◆ModuleProc名_指定ディレクトリ内のサブフォルダ一覧を取得してシートに書き出す
    '出力: -

    
'//処理実行
    Call callm1MexeFldrLstup
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            myXobjFldrPstdCellOUT As Object, _
            myXstrDirPathOUT As String, myXobjDirPstdCellOUT As Object)
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT
    Erase myZstrFldrNameOUT
    Erase myZstrFldrPathOUT
    Set myXobjFldrPstdCellOUT = Nothing
    myXstrDirPathOUT = Empty
    Set myXobjDirPstdCellOUT = Nothing
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonFldrCnt <= 0 Then Exit Sub
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    Set myXobjFldrPstdCellOUT = myXobjFldrPstdCell
    
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//S:エクセルシート上に記載されたフォルダパス一覧を取得
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:フォルダを選択してそのパスを取得してシートに書き出す
    Call snsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//S:指定ディレクトリ内のサブフォルダ一覧を取得してシートに書き出す
    Call snsProc3
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonFldrCnt = Empty: Erase myZobjFldr: Erase myZstrFldrName: Erase myZstrFldrPath
    Set myXobjFldrPstdCell = Nothing
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

'SnsP_エクセルシート上に記載されたフォルダパス一覧を取得する
Private Sub snsProc1()
    myXbisExitFlag = False
    
  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
        myXbisInStrOptn As Boolean
  Dim myXbisRowDrctn As Boolean
    
    Call m1Msub1ShtFldrLst.callProc( _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, _
            myXobjFldrPstdCell, _
            myXstrDirPath, myXobjDirPstdCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn, myXbisRowDrctn)
'    Debug.Print "データ: " & myXlonFldrCnt
    
    Set myXobjSrchSheet = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_フォルダを選択してそのパスを取得してシートに書き出す
Private Sub snsProc2()
    myXbisExitFlag = False
    
    If myXlonFldrCnt > 0 Then Exit Sub
    If PfncbisCheckFolderExist(myXstrDirPath) = True Then
        myXbisDirPstFlag = True
        Exit Sub
    End If

  Dim myXlonOutputOptn As Long, myXobjDirPstFrstCell As Object
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す
    
    If myXobjDirPstdCell Is Nothing Then
        myXlonOutputOptn = 0
        Set myXobjDirPstFrstCell = Nothing
    Else
        myXlonOutputOptn = 1
        Set myXobjDirPstFrstCell = myXobjDirPstdCell
    End If
    
  Dim myXlonDirSlctOptn As Long, _
        myXstrDfltFldrPath As String, myXlonIniView As Long, _
        myXbisExplrAdrsMsgOptn As Boolean
  
  Dim myXstrFldrPath As String, myXobjFldr As Object, _
        myXstrPrntPath As String, myXstrFldrName As String
    Call m1Msub2SlctFldrPathExtd.callProc( _
            myXbisDirPstFlag, _
            myXstrFldrPath, myXobjFldr, myXstrPrntPath, myXstrFldrName, _
            myXobjDirPstdCell, _
            myXlonDirSlctOptn, myXstrDfltFldrPath, myXlonIniView, myXbisExplrAdrsMsgOptn, _
            myXlonOutputOptn, myXobjDirPstFrstCell)
'    Debug.Print "データ: " & myXstrFldrPath
'    Debug.Print "データ: " & myXstrPrntPath
'    Debug.Print "データ: " & myXstrFldrName
    
    If myXstrFldrPath = "" Then GoTo ExitPath
    myXstrDirPath = myXstrFldrPath
    
    Set myXobjDirPstFrstCell = Nothing
    Set myXobjFldr = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_指定ディレクトリ内のサブフォルダ一覧を取得してシートに書き出す
Private Sub snsProc3()
    myXbisExitFlag = False
    
    If myXlonFldrCnt > 0 Then Exit Sub
    
  Dim myXlonOutputOptn As Long, myXobjFldrPstFrstCell As Object
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す
    
    If myXobjDirPstdCell Is Nothing Then
        myXlonOutputOptn = 1
        Set myXobjFldrPstFrstCell = Nothing
    Else
        Select Case myXbisDirPstFlag
        '//親フォルダパスがシートに記載されている場合
            Case True: myXlonOutputOptn = 2
            
        '//親フォルダパスがシートに記載されていない場合
            Case Else: myXlonOutputOptn = 1
        End Select
        Set myXobjFldrPstFrstCell = myXobjFldrPstdCell
    End If
    
  Dim myXbisNotOutFldrInfo As Boolean
    
  Dim myXbisCompFlag As Boolean
    Call m1Msub3SubFldrLstExtd.callProc( _
            myXbisCompFlag, _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, myXobjFldrPstdCell, _
            myXbisNotOutFldrInfo, myXstrDirPath, myXlonOutputOptn, myXobjFldrPstFrstCell)
'    Debug.Print "データ: " & myXlonFldrCnt
    
    Set myXobjFldrPstFrstCell = Nothing
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

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
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

'◆ModuleProc名_処理フォルダをリストアップする
Private Sub callm1MexeFldrLstup()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String, _
'        myXobjFldrPstdCell As Object, _
'        myXstrDirPath As String, myXobjDirPstdCell As Object
'    'myZobjFldr(k) : フォルダオブジェクト
'    'myZstrFldrName(k) : フォルダ名
'    'myZstrFldrPath(k) : フォルダパス
    Call m1MexeFldrLstup.callProc( _
            myXbisCmpltFlag, _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, _
            myXobjFldrPstdCell, myXstrDirPath, myXobjDirPstdCell)
    Call variablesOfm1MexeFldrLstup(myXlonFldrCnt, myZstrFldrPath)    'Debug.Print
End Sub
Private Sub variablesOfm1MexeFldrLstup( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//m1MexeFldrLstup内から出力した変数の内容確認
    Debug.Print "データ数: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myZvarField) To UBound(myZvarField)
        Debug.Print "データ" & k & ": " & myZvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInm1MexeFldrLstup()
'//m1MexeFldrLstupモジュールのモジュールメモリのリセット処理
    Call m1MexeFldrLstup.resetConstant
End Sub
