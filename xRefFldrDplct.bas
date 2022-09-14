Attribute VB_Name = "xRefFldrDplct"
'Includes CFldrDplct
'Includes PfnclonArrayDimension
'Includes PfncbisCheckFolderExist
'Includes PfixGetFolderNameInformationByFSO
'Includes PfncstrFolderPathReplaceParentBase
'Includes PfncbisCheckArrayDimension
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_フォルダを複製する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefFldrDplct"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  
'//入力制御信号
  
'//入力データ
  Private myXstrOrgFldrPath As String
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : 元情報
  Private myXstrSaveDirPath As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonRunInfoNo As Long, myXstrRunInfo As String
  Private myXlonDplctFldrCnt As Long, myZstrDplctFldrPath() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRunInfoNo = Empty: myXstrRunInfo = Empty
    myXlonDplctFldrCnt = Empty: Erase myZstrDplctFldrPath
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

'    myXlonOrgInfoCnt = 1
'    ReDim myZstrOrgInfo(myXlonOrgInfoCnt) As String
    
'//処理実行
    Call callxRefFldrDplct
    
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            ByVal myXstrOrgFldrPathIN As String, _
            ByVal myXlonOrgInfoCntIN As Long, ByRef myZstrOrgInfoIN() As String, _
            ByVal myXstrSaveDirPathIN As String)
    
'//入力変数を初期化
    myXstrOrgFldrPath = Empty
    myXlonOrgInfoCnt = Empty
    Erase myZstrOrgInfo
    myXstrSaveDirPath = Empty

'//入力変数を取り込み
    myXstrOrgFldrPath = myXstrOrgFldrPathIN
    If myXlonOrgInfoCntIN <= 0 Then Exit Sub
    myXlonOrgInfoCnt = myXlonOrgInfoCntIN
    myZstrOrgInfo() = myZstrOrgInfoIN()
    myXstrSaveDirPath = myXstrSaveDirPathIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
ExitPath:
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"    'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//指定フォルダの存在を確認
    If PfncbisCheckFolderExist(myXstrOrgFldrPath) = False Then GoTo ExitPath
    
'//指定フォルダの存在を確認
    If PfncbisCheckFolderExist(myXstrSaveDirPath) = False Then GoTo ExitPath
    
'//指定フォルダのフォルダ名情報を取得(FileSystemObject使用)
  Dim myXstrPrntPath As String, myXstrDirName As String
    Call PfixGetFolderNameInformationByFSO(myXstrPrntPath, myXstrDirName, myXstrOrgFldrPath)
    
'//C:情報一覧を順次実行
  Dim myXstrOrgPrnt As String, myXstrOrgBase As String
  Dim myXstrNewPrnt As String, myXstrNewBase As String
  Dim myXstrDplctFldrPath As String
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
        
    '//フォルダパス内のディレクトリパスorベース名を置換
        myXstrOrgPrnt = Empty: myXstrNewPrnt = Empty
        myXstrOrgBase = Empty: myXstrNewBase = Empty
        
        myXstrOrgPrnt = myXstrPrntPath
        myXstrNewPrnt = myXstrSaveDirPath
        myXstrDplctFldrPath = PfncstrFolderPathReplaceParentBase( _
                                myXstrOrgFldrPath, _
                                myXstrOrgPrnt, myXstrOrgBase, myXstrNewPrnt, myXstrNewBase)

        myXstrOrgPrnt = Empty: myXstrNewPrnt = Empty
        
        myXstrOrgBase = myXstrDirName
        myXstrNewBase = myXstrRunInfo
        myXstrDplctFldrPath = PfncstrFolderPathReplaceParentBase( _
                                myXstrDplctFldrPath, _
                                myXstrOrgPrnt, myXstrOrgBase, myXstrNewPrnt, myXstrNewBase)
        
        n = n + 1
        ReDim Preserve myZstrDplctFldrPath(n) As String
        myZstrDplctFldrPath(n) = myXstrDplctFldrPath
NextPath:
    Next k
    myXlonDplctFldrCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "7"    'PassFlag

'//ファイルを複製
    Call instCFldrDplct
    If myXbisExitFlag = True Then GoTo ExitPath
    
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

'RemP_保存した変数を取り出す
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
    
'//配列変数の次元数を取得
  Dim myXlonAryDmnsn As Long
    myXlonAryDmnsn = PfnclonArrayDimension(myZstrOrgInfo)
    
  Dim Li As Long, myXstrTmp As String
  
    On Error GoTo ExitPath
    Select Case myXlonAryDmnsn
        Case 1
            Li = LBound(myZstrOrgInfo): myXstrTmp = myZstrOrgInfo(Li)
            Exit Sub
        Case 2
            Li = LBound(myZstrOrgInfo, 1): myXstrTmp = myZstrOrgInfo(Li, Li)
        Case Else: GoTo ExitPath
    End Select
    On Error GoTo 0

  Dim myZstrOrgInfoINT() As String
    myZstrOrgInfoINT() = myZstrOrgInfo()
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
    i = UBound(myZstrOrgInfoINT, 1)
    ReDim myZstrOrgInfo(i) As String
    For i = LBound(myZstrOrgInfoINT) To UBound(myZstrOrgInfoINT)
        myZstrOrgInfo(i + Lo - Li) = myZstrOrgInfoINT(i)
    Next i
    
    Erase myZstrOrgInfoINT
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'CtrlP_
Private Sub ctrRunFiles()

'//C:情報一覧を順次実行
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim myXvarTmpInfo As Variant, k As Long: k = Li - 1
    For Each myXvarTmpPath In myZstrOrgInfo
        myXstrRunInfo = Empty
        k = k + 1: myXlonRunInfoNo = k
        myXstrRunInfo = CStr(myXvarTmpInfo)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next myXvarTmpInfo
    myXlonExeInfoCnt = n - Lo + 1
    
'//C:情報一覧を順次実行
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next k
    myXlonExeInfoCnt = n - Lo + 1
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
    
    myXstrOrgFldrPath = ThisWorkbook.Path & "\新しいフォルダー"
    
    myXstrSaveDirPath = ThisWorkbook.Path & "\try"
    
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_使用した変数を保存する
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

'◆ClassProc名_フォルダを複製する
Private Sub instCFldrDplct()
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsFldrDplct As CFldrDplct: Set myXinsFldrDplct = New CFldrDplct
    With myXinsFldrDplct
    '//クラス内変数への入力
        .letOrgFldrPath = myXstrOrgFldrPath
        .letAutoNaming = True
        If myXlonDplctFldrCnt <= 0 Then GoTo JumpPath
        .letDplctFldrCnt = myXlonDplctFldrCnt
        For i = 1 To myXlonDplctFldrCnt
            .letDplctFldrPathAry(i) = myZstrDplctFldrPath(i + L - 1)
        Next i
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXbisExitFlag = Not .getCmpltFlag
    End With
    Set myXinsFldrDplct = Nothing
JumpPath:
End Sub

'===============================================================================================
 
 '定型Ｆ_配列変数の次元数を取得する
Private Function PfnclonArrayDimension(ByRef myZvarOrgData As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarOrgData) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
End Function

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
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

 '定型Ｆ_フォルダパス内のディレクトリパスorベース名を置換する
Private Function PfncstrFolderPathReplaceParentBase( _
            ByVal myXstrOrgFldrPath As String, _
            ByVal myXstrOrgPrnt As String, ByVal myXstrOrgBase As String, _
            ByVal myXstrNewPrnt As String, ByVal myXstrNewBase As String) As String
    PfncstrFolderPathReplaceParentBase = Empty
    If myXstrOrgFldrPath = "" Then Exit Function
    If myXstrNewPrnt = "" And myXstrNewBase = "" Then Exit Function
  Dim myXstrNewFilePath As String
    If InStr(myXstrOrgFldrPath, myXstrOrgPrnt) > 0 And myXstrNewPrnt <> "" Then _
        myXstrNewFilePath = Replace(myXstrOrgFldrPath, myXstrOrgPrnt, myXstrNewPrnt)
    If InStr(myXstrOrgFldrPath, myXstrOrgBase) > 0 And myXstrNewBase <> "" Then _
        myXstrNewFilePath = Replace(myXstrOrgFldrPath, myXstrOrgBase, myXstrNewBase)
    PfncstrFolderPathReplaceParentBase = myXstrNewFilePath
End Function

 '定型Ｆ_配列変数の次元数が指定次元と一致するかをチェックする
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarOrgData As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarOrgData) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
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

'◆ModuleProc名_フォルダを複製する
Private Sub callxRefFldrDplct()
  Dim myXstrOrgFldrPath As String
  Dim myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
  Dim myXstrSaveDirPath As String
    'myZstrOrgInfo(i) : 元情報
    myXstrOrgFldrPath = ThisWorkbook.Path & "\新しいフォルダー"
    myXlonOrgInfoCnt = 2
    ReDim myZstrOrgInfo(2) As String
    myZstrOrgInfo(1) = "A"
    myZstrOrgInfo(2) = "B"
    myXstrSaveDirPath = ThisWorkbook.Path & "\try"
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonExeInfoCnt As Long, myZstrExeInfo() As String
'    'myZstrExeInfo(i) : 実行情報
    Call xRefFldrDplct.callProc( _
            myXbisCmpltFlag, _
            myXstrOrgFldrPath, myXlonOrgInfoCnt, myZstrOrgInfo, myXstrSaveDirPath)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub PforResetConstantInxRefFldrDplct()
'//xRefFldrDplctモジュールのモジュールメモリのリセット処理
    Call xRefFldrDplct.resetConstant
End Sub
