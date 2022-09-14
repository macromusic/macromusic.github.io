Attribute VB_Name = "xRefRunFiles"
'Includes PfncstrFileNameByFSO
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckArrayDimensionLength
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_複数ファイルに対して連続処理を実施する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRunFiles"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXlonExeFileCnt As Long, _
            myZstrExeFileName() As String, myZstrExeFilePath() As String
    'myZstrExeFileName(i) : 実行ファイル名
    'myZstrExeFilePath(i) : 実行ファイルパス
  
'//入力制御信号
  
'//入力データ
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
    'myZstrOrgFilePath(i) : 元ファイルパス
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
'  Private myZstrOrgFilePathINT() As String
  Private myXlonFileNo As Long, myXstrFileName As String, myXstrFilePath As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
'    Erase myZstrOrgFilePathINT
'    myXlonFileNo = Empty: myXstrFileName = Empty: myXstrFilePath = Empty
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_モジュールメモリのリセット
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
'    myZvarM(2, 1) = "meMvarField": myZvarM(2, 2) = Chr(34) & Chr(34)
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()

'//プログラム構成
    '入力: -
    '処理: -
    '出力: -

'    myXlonOrgFileCnt = 1
'    ReDim myZstrOrgFilePath(myXlonOrgFileCnt) As String
    
'//処理実行
    Call callxRefRunFiles
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, myXlonExeFileCntOUT As Long, _
            myZstrExeFileNameOUT() As String, myZstrExeFilePathOUT() As String, _
            ByVal myXlonOrgFileCntIN As Long, ByRef myZstrOrgFilePathIN() As String)
    
'//入力変数を初期化
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath

'//入力変数を取り込み
    If myXlonOrgFileCntIN <= 0 Then Exit Sub
    myXlonOrgFileCnt = myXlonOrgFileCntIN
    myZstrOrgFilePath() = myZstrOrgFilePathIN()
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    myXlonExeFileCntOUT = Empty
    Erase myZstrExeFileNameOUT: Erase myZstrExeFilePathOUT
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonExeFileCnt <= 0 Then Exit Sub
    myXlonExeFileCntOUT = myXlonExeFileCnt
    myZstrExeFileNameOUT() = myZstrExeFileName()
    myZstrExeFilePathOUT() = myZstrExeFilePath()
    
ExitPath:
End Sub

'PublicF_
Public Function fncbisCmpltFlag( _
                ByVal myXlonOrgFileCntIN As Long, _
                ByRef myZstrOrgFilePathIN() As String) As Boolean
    fncbisCmpltFlag = Empty
    
'//入力変数を初期化
    myXlonOrgFileCnt = Empty
    Erase myZstrOrgFilePath

'//入力変数を取り込み
    If myXlonOrgFileCntIN <= 0 Then Exit Sub
    myXlonOrgFileCnt = myXlonOrgFileCntIN
    myZstrOrgFilePath = myZstrOrgFilePathIN
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Function
    
    fncbisCmpltFlag = myXbisCmpltFlag
    
ExitPath:
End Function

'-----------------------------------------------------------------------------------------------
'Control  : ユーザから入力を受け取ってその内容に応じてSense、Process、Runを制御する
'Sense    : Processで実行する演算処理用のデータを取得する
'Process  : Senseで取得したデータを使用して演算処理をする
'Run      : Processの処理結果を受けて画面表示などの出力処理をする
'Remember : 記録した内容を必要に応じて取り出して処理に活用する
'Record   : Sense、Process、Runで実行したプログラムで重要な内容を記録する
'-----------------------------------------------------------------------------------------------

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"    'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:Loop前の情報取得処理
    Call snsProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"    'PassFlag
    
'//P:Loop前の情報加工処理
    Call prsProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"    'PassFlag
    
'//C:ファイルリストを順次実行
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim k As Long
    For k = LBound(myZstrOrgFilePath) To UBound(myZstrOrgFilePath)
        myXstrFilePath = Empty: myXstrFileName = Empty
        myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
 
    '//S:各ファイルのデータ取得処理
        Call snsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "4-" & k   'PassFlag
 
    '//P:各ファイルのデータ加工処理
        Call prsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "5-" & k   'PassFlag
            
    '//Run:各ファイルのデータ出力処理
        Call runProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "6-" & k   'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeFileName(n) As String
        ReDim Preserve myZstrExeFilePath(n) As String
        myZstrExeFileName(n) = myXstrFileName
        myZstrExeFilePath(n) = myXstrFilePath
NextPath:
    Next k
    myXlonExeFileCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "7"    'PassFlag
    
'//P:Loop後の加工処理
    Call prsProcAfterLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "8"    'PassFlag

'//Run:ファイナライズ処理
    Call runFinalize
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "9"   'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
'    myXbisCmpltFlag = False
'    myXlonExeFileCnt = Empty
'    Erase myZstrExeFileName: Erase myZstrExeFilePath
End Sub

'RemP_保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
'    If myXvarFieldIN = Empty Then myXvarFieldIN = meMvarField
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'  Dim Li As Long, myXstrTmp As String
'    On Error GoTo ExitPath
'    Li = LBound(myZstrOrgFilePath): myXstrTmp = myZstrOrgFilePath(Li)
'    On Error GoTo 0
    
'//入力配列変数を内部配列変数に入れ替える
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim Li As Long, Ui As Long, i As Long
'    On Error GoTo ExitPath
'    Li = LBound(myZstrOrgFilePath): Ui = UBound(myZstrOrgFilePath)
'    i = Ui + Lo - Li: ReDim myZstrOrgFilePathINT(i) As String
'    For i = Li To Ui
'        myZstrOrgFilePathINT(i + Lo - Li, j + Lo - Li) = myZstrOrgFilePath(i, j)
'    Next i
'    On Error GoTo 0
    
'//入力配列変数の内容を確認
'    If PfncbisCheckArrayDimension(myZstrOrgFilePath, 1) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'CtrlP_
Private Sub ctrRunFiles()

'//C:ファイルリストを順次実行
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim myXvarTmpPath As Variant, k As Long: k = Li - 1
    For Each myXvarTmpPath In myZstrOrgFilePath
        myXstrFilePath = Empty: myXstrFileName = Empty
        k = k + 1: myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
        'XarbProgCode
        n = n + 1
        ReDim Preserve myZstrExeOrgFileName(n) As String
        ReDim Preserve myZstrExeOrgFilePath(n) As String
        myZstrExeOrgFileName(n) = myXstrFileName
        myZstrExeOrgFilePath(n) = myXstrFilePath
NextPath:
    Next myXvarTmpPath
    myXlonExeOrgFileCnt = n - Lo + 1
    
'//C:ファイルリストを順次実行
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim k As Long
    For k = LBound(myZstrOrgFilePath) To UBound(myZstrOrgFilePath)
        myXstrFilePath = Empty: myXstrFileName = Empty
        myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
        'XarbProgCode
        n = n + 1
        ReDim Preserve myZstrExeFileName(n) As String
        ReDim Preserve myZstrExeFilePath(n) As String
        myZstrExeFileName(n) = myXstrFileName
        myZstrExeFilePath(n) = myXstrFilePath
NextPath:
    Next k
    myXlonExeFileCnt = n - Lo + 1
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
End Sub

'SnsP_Loop前の情報取得処理
Private Sub snsProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop前の情報加工処理
Private Sub prsProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_各ファイルのデータ取得処理
Private Sub snsProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_各ファイルのデータ加工処理
Private Sub prsProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_各ファイルのデータ出力処理
Private Sub runProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "6-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop後の加工処理
Private Sub prsProcAfterLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "8-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_ファイナライズ処理
Private Sub runFinalize()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "9-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
'    If myXlonExeFileCnt <= 0 Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeFileName, 1) = False Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeFilePath, 1) = False Then GoTo ExitPath
    
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
'    myZvarM(1, 1) = "meMvarField"
'    myZvarM(1, 2) = myXvarField
    
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

 '定型Ｆ_指定ファイルのファイル名を取得する(FileSystemObject使用)
Private Function PfncstrFileNameByFSO(ByVal myXstrFilePath As String) As String
    PfncstrFileNameByFSO = Empty
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExist As Boolean: myXbisFileExist = myXobjFSO.FileExists(myXstrFilePath)
    If myXbisFileExist = False Then Exit Function
    PfncstrFileNameByFSO = myXobjFSO.GetFileName(myXstrFilePath)
    Set myXobjFSO = Nothing
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

 '定型Ｆ_配列変数の次元数と配列長が指定値を満足するかをチェックする
Private Function PfncbisCheckArrayDimensionLength( _
            ByRef myZvarOrgData As Variant, ByVal myXlonChckAryDmnsn As Long, _
            ByRef myXlonChckAryLen() As Long) As Boolean
'myXlonChckAryDmnsn  : 配列の次元数の指定値
'myXlonChckAryLen(i) : i次元目の配列長の指定値
'myXlonChckAryLen(i) = 0 : 配列長のチェックを実施しない
    PfncbisCheckArrayDimensionLength = False
    If myXlonChckAryDmnsn <= 0 Then Exit Function
  Dim Li As Long, Ui As Long, myXlonChckAryLenCnt As Long
    On Error Resume Next
    Li = LBound(myXlonChckAryLen): Ui = UBound(myXlonChckAryLen)
    If Err.Number = 9 Then Exit Function
    On Error GoTo 0
    myXlonChckAryLenCnt = Ui - Li + 1
    If myXlonChckAryLenCnt <= 0 Then Exit Function
  Dim i As Long
    For i = LBound(myXlonChckAryLen) To UBound(myXlonChckAryLen)
        If myXlonChckAryLen(i) <= 0 Then Exit Function
    Next i
'//配列であることを確認
    If IsArray(myZvarOrgData) = False Then Exit Function
'//配列が空でないことを確認
  Dim myXlonTmp As Long
    On Error Resume Next
    myXlonTmp = UBound(myZvarOrgData) - LBound(myZvarOrgData) + 1
    If Err.Number = 9 Then Exit Function
    On Error GoTo 0
    If myXlonTmp <= 0 Then Exit Function
'//配列の次元数を取得
  Dim myXlonAryDmnsn As Long, myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    myXlonAryDmnsn = k - 1
    If myXlonAryDmnsn <> myXlonChckAryDmnsn Then Exit Function
    If myXlonAryDmnsn <> myXlonChckAryLenCnt Then Exit Function
'//配列の最小添字と最大添字を取得
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    i = myXlonAryDmnsn + L - 1
  Dim myZlonAryLBnd() As Long: ReDim myZlonAryLBnd(i) As Long
  Dim myZlonAryUBnd() As Long: ReDim myZlonAryUBnd(i) As Long
    k = 0
    For i = LBound(myZlonAryLBnd) To UBound(myZlonAryLBnd)
        k = k + 1
        myZlonAryLBnd(i) = LBound(myZvarOrgData, k)
        myZlonAryUBnd(i) = UBound(myZvarOrgData, k)
    Next i
'//配列長を取得
    i = myXlonAryDmnsn + L - 1
  Dim myZlonAryLen() As Long: ReDim myZlonAryLen(i) As Long
    For i = LBound(myZlonAryLen) To UBound(myZlonAryLen)
        myZlonAryLen(i) = myZlonAryUBnd(i) - myZlonAryLBnd(i) + 1
    Next i
'//次元数と配列長をチェック
    For i = LBound(myZlonAryLen) To UBound(myZlonAryLen)
        If myXlonChckAryLen(i + Li - L) <> 0 Then _
            If myZlonAryLen(i) <> myXlonChckAryLen(i + Li - L) Then Exit Function
    Next i
    PfncbisCheckArrayDimensionLength = True
    Erase myZlonAryLBnd: Erase myZlonAryUBnd: Erase myZlonAryLen
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

'◆ModuleProc名_複数ファイルに対して連続処理を実施する
Private Sub callxRefRunFiles()
  Dim myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
    'myZstrOrgFilePath(i) : 元ファイルパス
    myXlonOrgFileCnt = XarbLong
    ReDim myZstrOrgFilePath(1) As String
    myZstrOrgFilePath(1) = XarbString
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonExeFileCnt As Long, _
'        myZstrExeFileName() As String, myZstrExeFilePath() As String
'    'myZstrExeFileName(i) : 実行ファイル名
'    'myZstrExeFilePath(i) : 実行ファイルパス
    Call xRefRunFiles.callProc( _
            myXbisCmpltFlag, _
            myXlonExeFileCnt, myZstrExeFileName, myZstrExeFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath)
    Call variablesOfxRefRunFiles(myXlonExeFileCnt, myZstrExeFileName)    'Debug.Print
End Sub
Private Sub variablesOfxRefRunFiles( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefRunFiles内から出力した変数の内容確認
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
Private Sub resetConstantInxRefRunFiles()
'//xRefRunFilesモジュールのモジュールメモリのリセット処理
    Call xRefRunFiles.resetConstant
End Sub
