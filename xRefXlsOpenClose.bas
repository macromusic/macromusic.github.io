Attribute VB_Name = "xRefXlsOpenClose"
'Includes CXlsOpen
'Includes CXlsClose
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckFileExist
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_エクセルブックを開閉して処理する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefXlsOpenClose"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXobjOpndBook As Object
  
'//入力制御信号
  
'//入力データ
  Private myXbisOpnRdOnly As Boolean
  Private myXstrOpnFullName As String
  
  Private myXbisSaveON As Boolean
  Private myXstrSaveBkName As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  Private myXbisErrFlag As Boolean
  
'//モジュール内変数_データ
  Private myXstrCloseFullName As String
  Private myXbisBkOpnd As Boolean
    'myXbisBkOpnd = True  : 指定エクセルブックが開いている
    'myXbisBkOpnd = False : 指定エクセルブックが開いていない
  Private myXbisBkRdOnly As Boolean
    'myXbisBkRdOnly = True  : 指定エクセルブックが読み取り専用で開いている
    'myXbisBkRdOnly = False : 指定エクセルブックが読み取り専用では開いていない

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrCloseFullName = Empty
    myXbisBkOpnd = False: myXbisBkRdOnly = False
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
    Call callxRefXlsOpenClose
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXobjOpndBookOUT As Object, _
            ByVal myXbisOpnRdOnlyIN As Boolean, ByVal myXstrOpnFullNameIN As String, _
            ByVal myXbisSaveONIN As Boolean, ByVal myXstrSaveBkNameIN As String)
    
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    
'//入力変数を初期化
    myXbisOpnRdOnly = False
    myXstrOpnFullName = Empty
    
    myXbisSaveON = False
    myXstrSaveBkName = Empty

'//入力変数を取り込み
    myXbisOpnRdOnly = myXbisOpnRdOnlyIN
    myXstrOpnFullName = myXstrOpnFullNameIN
    
    myXbisSaveON = myXbisSaveONIN
    myXstrSaveBkName = myXstrSaveBkNameIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    Set myXobjOpndBookOUT = Nothing
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    Set myXobjOpndBookOUT = myXobjOpndBook

ExitPath:
End Sub

'PublicF_
Public Function fncobjOpenedBook( _
            ByVal myXbisOpnRdOnlyIN As Boolean, ByVal myXstrOpnFullNameIN As String, _
            ByVal myXbisSaveONIN As Boolean, ByVal myXstrSaveBkNameIN As String) As Object
    Set fncobjOpenedBook = Nothing
    
'//入力変数を初期化
    myXbisOpnRdOnly = False
    myXstrOpnFullName = Empty
    
    myXbisSaveON = False
    myXstrSaveBkName = Empty

'//入力変数を取り込み
    myXbisOpnRdOnly = myXbisOpnRdOnlyIN
    myXstrOpnFullName = myXstrOpnFullNameIN
    
    myXbisSaveON = myXbisSaveONIN
    myXstrSaveBkName = myXstrSaveBkNameIN
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Function
    
    Set fncobjOpenedBook = myXobjOpndBook
    
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:
    Call setControlVariables1
    Call setControlVariables2
    
'//S:エクセルブックを開く
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:エクセルブックを閉じる
    Call prsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
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
    Set myXobjOpndBook = Nothing
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
    
'//指定ファイルの存在を確認
    If PfncbisCheckFileExist(myXstrOpnFullName) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables1()
    
    myXbisOpnRdOnly = True
    'myXbisOpnRdOnly = False : 指定エクセルブックを読み取り専用にせずに開く
    'myXbisOpnRdOnly = True  : 指定エクセルブックを読み取り専用で開く
    
    myXstrOpnFullName = ThisWorkbook.Path & "\新規 Microsoft Excel ワークシート.xlsx"
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    myXbisSaveON = False
    'myXbisSaveON = False : 閉じる前に保存しない
    'myXbisSaveON = True  : 閉じる前に保存する
    
    myXstrSaveBkName = ""
    
End Sub

'SnsP_エクセルブックを開く
Private Sub snsProc1()
    myXbisExitFlag = False

    Call instCXlsOpen
    If myXobjOpndBook Is Nothing Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_
Private Sub prsProc1()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3-1"   'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_エクセルブックを閉じる
Private Sub prsProc2()
    myXbisExitFlag = False

    myXstrCloseFullName = myXstrOpnFullName
    
    Call instCXlsClose
    If myXbisErrFlag = True Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_
Private Sub runProc()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5-1"   'PassFlag
    
    'XarbProgCode
    
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

'モジュール内Ｐ_
Private Sub MsubProc()
End Sub

'モジュール内Ｆ_
Private Function MfncFunc() As Variant
End Function

'===============================================================================================

'◆ClassProc名_エクセルブックを開く
Private Sub instCXlsOpen()
  Dim myXinsXlsOpen As CXlsOpen: Set myXinsXlsOpen = New CXlsOpen
    With myXinsXlsOpen
    '//クラス内変数への入力
        .letOpnRdOnly = myXbisOpnRdOnly
        .letOpnFullName = myXstrOpnFullName
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXbisCmpltFlag = .getCmpltFlag
        Set myXobjOpndBook = .getOpndBook
        Set myXobjOpndBook = .fncobjOpenedBook
    End With
    Set myXinsXlsOpen = Nothing
End Sub

'◆ClassProc名_エクセルブックを閉じる
Private Sub instCXlsClose()
  Dim myXinsXlsClose As CXlsClose: Set myXinsXlsClose = New CXlsClose
    With myXinsXlsClose
    '//クラス内変数への入力
        .letCloseFullName = myXstrCloseFullName
        .letSaveON = myXbisSaveON
        .letSaveBkName = myXstrSaveBkName
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXbisErrFlag = Not .getCmpltFlag
        myXbisBkOpnd = .getBkOpnd
        myXbisBkRdOnly = .getBkRdOnly
    End With
    Set myXinsXlsClose = Nothing
End Sub

'===============================================================================================

 '定型Ｐ_
Private Sub PfixProc()
End Sub

 '定型Ｆ_
Private Function PfncFunc() As Variant
End Function

 '定型Ｆ_配列変数の次元数が指定次元と一致するかをチェックする
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarDataAry As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarDataAry) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarDataAry, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
End Function
 
 '定型Ｆ_指定ファイルの存在を確認する
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
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
'    myXbisOpnRdOnly = True
'    'myXbisOpnRdOnly = False : 指定エクセルブックを読み取り専用にせずに開く
'    'myXbisOpnRdOnly = True  : 指定エクセルブックを読み取り専用で開く
'    myXstrOpnFullName = ""
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    myXbisSaveON = False
'    'myXbisSaveON = False : 閉じる前に保存しない
'    'myXbisSaveON = True  : 閉じる前に保存する
'    myXstrSaveBkName = ""
'End Sub
'◆ModuleProc名_エクセルブックを開閉して処理する
Private Sub callxRefXlsOpenClose()
  Dim myXbisOpnRdOnly As Boolean, myXstrOpnFullName As String
    'myXbisOpnRdOnly = False : 指定エクセルブックを読み取り専用にせずに開く
    'myXbisOpnRdOnly = True  : 指定エクセルブックを読み取り専用で開く
    myXstrOpnFullName = ThisWorkbook.Path & "\新規 Microsoft Excel ワークシート.xlsx"
  Dim myXbisSaveON As Boolean, myXstrSaveBkName As String
    'myXbisSaveON = False : 閉じる前に保存しない
    'myXbisSaveON = True  : 閉じる前に保存する
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjOpndBook As Object
    Call xRefXlsOpenClose.callProc( _
            myXbisCmpltFlag, myXobjOpndBook, _
            myXbisOpnRdOnly, myXstrOpnFullName, _
            myXbisSaveON, myXstrSaveBkName)
'    Set myXobjOpndBook = xRefXlsOpenClose.fncobjOpenedBook( _
'            myXbisOpnRdOnly, myXstrOpnFullName, _
'            myXbisSaveON, myXstrSaveBkName)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefXlsOpenClose()
'//xRefXlsOpenCloseモジュールのモジュールメモリのリセット処理
    Call xRefXlsOpenClose.resetConstant
End Sub
