VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckArrayDimension
'Includes PincCreateTextFileWithAutoNamingAndHierarchyCheck
'Includes PincFileAutoNamingAndHierarchyCheck
'Includes PfnclonCountFileLayer
'Includes PfncbisCheckFileExist
'Includes PfixGetFileNameInformation
'Includes PfncstrFileAutoNaming
'Includes PincCreateFolderWithAutoNamingAndHierarchyCheck
'Includes PincFolderAutoNamingAndHierarchyCheck
'Includes PfnclonCountFolderLayer
'Includes PfncbisCheckFolderExist
'Includes PfncstrFolderAutoNaming
'Includes PfixGetFolderNameInformation
'Includes PfixCreateFolderWithHierarchyCheck
'Includes PfixCreateFolder
'Includes PfixCreateTextFile
'Includes PfixGetFileNameInformationByFSO

Option Explicit
Option Base 1

'◆ClassProc名_テキストファイルを新規作成する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CFileCreate"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonExeFileCnt As Long, _
            myZstrExeFilePath() As String, myZstrExeFileName() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisAutoNaming As Boolean
  
'//PropertySetLet_入力データ
  Private myXlonCreateFileCnt As Long, myZstrCreateFilePath() As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXstrNewFilePath As String
  Private myXstrExeFilePath As String, myXstrExeFileName As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrNewFilePath = Empty
    myXstrExeFilePath = Empty: myXstrExeFileName = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getCmpltFlag() As Boolean
    getCmpltFlag = myXbisCmpltFlag
End Property

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getExeFileCnt() As Long
    getExeFileCnt = myXlonExeFileCnt
End Property

Public Property Get getExeFilePathAry(ByVal i As Long) As String
    getExeFilePathAry = myZstrExeFilePath(i)
'myZstrExeFilePath(i) : 実行したファイルパス
End Property

Public Property Get getExeFileNameAry(ByVal i As Long) As String
    getExeFileNameAry = myZstrExeFileName(i)
'myZstrExeFileName(i) : 実行したファイル名
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonExeFileCnt = Empty: Erase myZstrExeFilePath: Erase myZstrExeFileName
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letCreateFileCnt(ByVal RHS As Long)
    myXlonCreateFileCnt = RHS
End Property

Public Property Let letCreateFilePathAry(ByRef RHS() As String)
'myZstrCreateFilePath(i) : 元データ配列
    myZstrCreateFilePath() = RHS()
End Property

Public Property Let letAutoNaming(ByVal RHS As Boolean)
    myXbisAutoNaming = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If PfncbisCheckArrayDimension(myZstrCreateFilePath, 1) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    Call ctrProc
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"  'PassFlag
    
'//C:ファイルリストを順次実行
  Dim myXvarTmpPath As Variant, n As Long: n = 0
    For Each myXvarTmpPath In myZstrCreateFilePath
        myXstrNewFilePath = CStr(myXvarTmpPath)
 
    '//P:各ファイルのデータ加工処理
        Call prsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "2-" & k  'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeFilePath(n) As String
        ReDim Preserve myZstrExeFileName(n) As String
        myZstrExeFilePath(n) = myXstrExeFilePath
        myZstrExeFileName(n) = myXstrExeFileName
NextPath:
    Next
    myXlonExeFileCnt = n
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

'PrcsP_各ファイルのデータ加工処理
Private Sub prsProcForLoop()
    myXbisExitFlag = False
 
'//Textファイルを新規作成
    Call PincCreateTextFileWithAutoNamingAndHierarchyCheck( _
            myXstrExeFilePath, _
            myXstrNewFilePath, , myXbisAutoNaming)
    If myXstrExeFilePath = "" Then GoTo ExitPath
    
'//指定ファイルのファイル名情報を取得(FileSystemObject使用)
  Dim myXstrPrntPath As String, myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformationByFSO( _
            myXstrPrntPath, myXstrExeFileName, myXstrBaseName, myXstrExtsn, _
            myXstrExeFilePath)
    If myXstrExeFileName = "" Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

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

 '定型Ｐ_Textファイルを新規作成する(ファイル名自動生成＆階層チェック機能付き)
Private Sub PincCreateTextFileWithAutoNamingAndHierarchyCheck( _
            myXstrFilePathOUT As String, _
            ByVal myXstrFilePathIN As String, _
            Optional ByVal coXlonNGLayer As Long = 2, _
            Optional ByVal coXbisAutoNaming As Boolean = False)
'Includes PincFileAutoNamingAndHierarchyCheck
'Includes PfixCreateTextFile
'coXlonNGLayer >= 1  : ファイルの必要階層数
'coXbisAutoNaming = False  : 同一名ファイルがあれば終了
'coXbisAutoNaming = True   : 同一名ファイルがあればファイル名を自動生成
    myXstrFilePathOUT = Empty
'//新規ファイルのファイル名自動生成＆階層チェック
  Dim myXbisFileExistFlag As Boolean
    Call PincFileAutoNamingAndHierarchyCheck( _
            myXbisFileExistFlag, myXstrFilePathOUT, _
            myXstrFilePathIN, coXlonNGLayer, coXbisAutoNaming)
    If myXstrFilePathOUT = "" Then GoTo ExitPath
'//Txtファイルを新規作成
  Dim myXbisExitFlag As Boolean
    Call PfixCreateTextFile(myXbisExitFlag, myXstrFilePathOUT)
    If myXbisExitFlag = True Then GoTo ExitPath
    Exit Sub
ExitPath:
    myXstrFilePathOUT = ""
End Sub

 '定型Ｐ_ファイル名自動生成＆階層チェック
Private Sub PincFileAutoNamingAndHierarchyCheck( _
            myXbisFileExistFlag As Boolean, myXstrFilePathOUT As String, _
            ByVal myXstrFilePathIN As String, _
            Optional ByVal coXlonNGLayer As Long = 2, _
            Optional ByVal coXbisAutoNaming As Boolean = False)
'Includes PfnclonCountFileLayer
'Includes PfncbisCheckFileExist
'Includes PfixGetFileNameInformation
'Includes PfncstrFileAutoNaming
'Includes PincCreateFolderWithAutoNamingAndHierarchyCheck
'coXlonNGLayer >= 1  : ファイルの必要階層数
'coXbisAutoNaming = False  : 同一名ファイルがあれば終了
'coXbisAutoNaming = True   : 同一名ファイルがあればファイル名を自動生成
    myXbisFileExistFlag = False: myXstrFilePathOUT = Empty
'//対象ファイルが指定階層より下層でなければ中止
    If coXlonNGLayer <= 0 Then GoTo ExitPath
    If PfnclonCountFileLayer(myXstrFilePathIN) <= coXlonNGLayer Then GoTo ExitPath
'//対象ファイルと同一名ファイルが存在するか確認
    myXbisFileExistFlag = PfncbisCheckFileExist(myXstrFilePathIN)
'//対象ファイルのファイル名情報を取得
  Dim myXstrPrntPath As String, myXstrFileName As String, _
        myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformation( _
            myXstrPrntPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrFilePathIN)
    If myXbisFileExistFlag = True Then
    '//同一名ファイルが存在する場合
    '//ファイル名を自動生成
        If coXbisAutoNaming = False Then GoTo ExitPath
        myXstrFilePathOUT = PfncstrFileAutoNaming( _
                                myXstrPrntPath, myXstrBaseName, myXstrExtsn)
    Else
    '//同一名ファイルが存在しない場合
    '//フォルダ階層をチェック
        If PfnclonCountFileLayer(myXstrPrntPath) <= 0 Then GoTo ExitPath
      Dim myXstrDirPath As String
        Call PincCreateFolderWithAutoNamingAndHierarchyCheck( _
                myXstrDirPath, _
                myXstrPrntPath, coXlonNGLayer - 1, False)
        If myXstrDirPath = "" Then GoTo ExitPath
        myXstrFilePathOUT = myXstrDirPath & "\" & myXstrBaseName & "." & myXstrExtsn
    End If
    Exit Sub
ExitPath:
    myXstrFilePathOUT = ""
End Sub

 '定型Ｆ_指定ファイルの階層位置を取得する
Private Function PfnclonCountFileLayer(ByVal myXstrFilePath As String) As Long
    PfnclonCountFileLayer = Empty
    If myXstrFilePath = "" Then Exit Function
    If Right(myXstrFilePath, 1) = "\" Then Exit Function
  Dim myXstrTmpPath As String: myXstrTmpPath = myXstrFilePath
    If Left(myXstrTmpPath, 2) = "\\" Then myXstrTmpPath = Mid(myXstrTmpPath, 3)
  Dim i As Long, n As Long: n = 0
    For i = Len(myXstrTmpPath) To 1 Step -1
        If Mid(myXstrTmpPath, i, 1) = "\" Then n = n + 1
    Next i
    PfnclonCountFileLayer = n
End Function

 '定型Ｆ_指定ファイルの存在を確認する
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_指定ファイルのファイル名情報を取得する
Private Sub PfixGetFileNameInformation( _
            myXstrPrntPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String, _
            ByVal myXstrFilePath As String)
    myXstrPrntPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    If myXstrFilePath = "" Then Exit Sub
  Dim i As Long, myXlonInStrRev As Long: myXlonInStrRev = 0
    For i = Len(myXstrFilePath) To 1 Step -1
        If Mid(myXstrFilePath, i, 1) = "\" Then
            myXlonInStrRev = i: Exit For
        End If
    Next i
    If myXlonInStrRev = 0 Then
        myXstrFileName = myXstrFilePath
    ElseIf myXlonInStrRev = 1 Then
        myXstrFileName = Mid(myXstrFilePath, myXlonInStrRev + 1)
    ElseIf myXlonInStrRev = Len(myXstrFilePath) Then
        myXstrPrntPath = Left(myXstrFilePath, myXlonInStrRev - 1)
    Else
        myXstrPrntPath = Left(myXstrFilePath, myXlonInStrRev - 1)
        myXstrFileName = Mid(myXstrFilePath, myXlonInStrRev + 1)
    End If
    myXlonInStrRev = 0
    For i = Len(myXstrFileName) To 1 Step -1
        If Mid(myXstrFileName, i, 1) = "." Then
            myXlonInStrRev = i: Exit For
        End If
    Next i
    If myXlonInStrRev = 0 Then
        myXstrBaseName = myXstrFileName
    ElseIf myXlonInStrRev = 1 Then
        myXstrExtsn = Mid(myXstrFileName, myXlonInStrRev + 1)
    ElseIf myXlonInStrRev = Len(myXstrFileName) Then
        myXstrBaseName = Left(myXstrFileName, myXlonInStrRev - 1)
    Else
        myXstrBaseName = Left(myXstrFileName, myXlonInStrRev - 1)
        myXstrExtsn = Mid(myXstrFileName, myXlonInStrRev + 1)
    End If
End Sub

 '定型Ｆ_ファイル名を自動生成する
Private Function PfncstrFileAutoNaming( _
            ByVal myXstrPrntPath As String, ByVal myXstrBaseName As String, _
            ByVal myXstrExtsn As String) As String
    PfncstrFileAutoNaming = Empty
    If myXstrPrntPath = "" Or myXstrExtsn = "" Then Exit Function
  Dim myXstrTmpBase As String
    myXstrTmpBase = myXstrBaseName
    If myXstrBaseName = "" Then _
        myXstrTmpBase = myXstrBaseName & "_" & Format(Now, "YYYYMMDDHHmmSS")
  Dim myXstrFilePath As String
    myXstrFilePath = myXstrPrntPath & "\" & myXstrTmpBase & "." & myXstrExtsn
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExistFlag As Boolean
    myXbisFileExistFlag = myXobjFSO.FileExists(myXstrFilePath)
    Do While myXbisFileExistFlag = True
        myXstrTmpBase = myXstrBaseName & "_" & Format(Now, "YYYYMMDDHHmmSS")
        myXstrFilePath = myXstrPrntPath & "\" & myXstrTmpBase & "." & myXstrExtsn
        myXbisFileExistFlag = myXobjFSO.FileExists(myXstrFilePath)
    Loop
    PfncstrFileAutoNaming = myXstrFilePath
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_フォルダを新規作成する(フォルダ名自動生成＆階層チェック機能付き)
Private Sub PincCreateFolderWithAutoNamingAndHierarchyCheck( _
            myXstrDirPathOUT As String, _
            ByVal myXstrDirPathIN As String, _
            Optional ByVal coXlonNGLayer As Long = 1, _
            Optional ByVal coXbisAutoNaming As Boolean = False)
'Includes PincFolderAutoNamingAndHierarchyCheck
'Includes PfixCreateFolder
'coXlonNGLayer >= 1  : フォルダの必要階層数
'coXbisAutoNaming = False  : 同一名フォルダがあれば終了
'coXbisAutoNaming = True   : 同一名フォルダがあればフォルダ名を自動生成
    myXstrDirPathOUT = Empty
'//新規フォルダのフォルダ名自動生成＆階層チェック
  Dim myXbisDirExistFlag As Boolean
    Call PincFolderAutoNamingAndHierarchyCheck( _
            myXbisDirExistFlag, myXstrDirPathOUT, _
            myXstrDirPathIN, coXlonNGLayer, coXbisAutoNaming)
    If myXstrDirPathOUT = "" Then GoTo ExitPath
    If myXbisDirExistFlag = True Then _
        If myXstrDirPathOUT = myXstrDirPathIN Then Exit Sub
'//フォルダを新規作成
  Dim myXbisExitFlag As Boolean
    Call PfixCreateFolder(myXbisExitFlag, myXstrDirPathOUT)
    If myXbisExitFlag = True Then GoTo ExitPath
    Exit Sub
ExitPath:
    myXstrDirPathOUT = ""
End Sub

 '定型Ｐ_フォルダ名自動生成＆階層チェック
Private Sub PincFolderAutoNamingAndHierarchyCheck( _
            myXbisDirExistFlag As Boolean, myXstrDirPathOUT As String, _
            ByVal myXstrDirPathIN As String, _
            Optional ByVal coXlonNGLayer As Long = 1, _
            Optional ByVal coXbisAutoNaming As Boolean = False)
'Includes PfnclonCountFolderLayer
'Includes PfncbisCheckFolderExist
'Includes PfncstrFolderAutoNaming
'Includes PfixGetFolderNameInformation
'Includes PfixCreateFolderWithHierarchyCheck
'coXlonNGLayer >= 1  : フォルダの必要階層数
'coXbisAutoNaming = False  : 同一名フォルダがあれば終了
'coXbisAutoNaming = True   : 同一名フォルダがあればフォルダ名を自動生成
    myXbisDirExistFlag = False: myXstrDirPathOUT = Empty
'//対象フォルダが指定階層より下層でなければ中止
    If coXlonNGLayer <= 0 Then GoTo ExitPath
    If PfnclonCountFolderLayer(myXstrDirPathIN) <= coXlonNGLayer Then GoTo ExitPath
'//対象フォルダと同一名フォルダが存在するか確認
    myXbisDirExistFlag = PfncbisCheckFolderExist(myXstrDirPathIN)
    If myXbisDirExistFlag = True Then
    '//同一名フォルダが存在する場合
    '//フォルダ名を自動生成
        Select Case coXbisAutoNaming
            Case True: myXstrDirPathOUT = PfncstrFolderAutoNaming(myXstrDirPathIN)
            Case Else: myXstrDirPathOUT = myXstrDirPathIN: Exit Sub
        End Select
    Else
    '//同一名フォルダが存在しない場合
    '//フォルダ階層をチェック
      Dim myXstrPrntPath As String, myXstrDirName As String
        Call PfixGetFolderNameInformation(myXstrPrntPath, myXstrDirName, myXstrDirPathIN)
      Dim myXstrBaseDirPath As String, myXstrRltvDirPath As String
        If Left(myXstrPrntPath, 2) = "\\" Then
          Dim myXstrTmpPath As String: myXstrTmpPath = Mid(myXstrPrntPath, 3)
            myXstrBaseDirPath = Left(myXstrPrntPath, InStr(myXstrTmpPath, "\") + 2 - 1)
        Else
            myXstrBaseDirPath = Left(myXstrPrntPath, InStr(myXstrPrntPath, "\") - 1)
        End If
        myXstrRltvDirPath = Mid(myXstrPrntPath, Len(myXstrBaseDirPath) + 2)
      Dim myXbisExitFlag As Boolean
        Call PfixCreateFolderWithHierarchyCheck(myXbisExitFlag, _
                myXstrBaseDirPath, myXstrRltvDirPath)
        If myXbisExitFlag = True Then GoTo ExitPath
        myXstrDirPathOUT = myXstrDirPathIN
    End If
    Exit Sub
ExitPath:
    myXstrDirPathOUT = ""
End Sub

 '定型Ｆ_指定フォルダの階層位置を取得する
Private Function PfnclonCountFolderLayer(ByVal myXstrDirPath As String) As Long
    PfnclonCountFolderLayer = Empty
    If myXstrDirPath = "" Then Exit Function
    If Right(myXstrDirPath, 1) = "\" Then Exit Function
  Dim myXstrTmpPath As String: myXstrTmpPath = myXstrDirPath
    If Left(myXstrTmpPath, 2) = "\\" Then myXstrTmpPath = Mid(myXstrTmpPath, 3)
  Dim i As Long, n As Long: n = 0
    For i = Len(myXstrTmpPath) To 1 Step -1
        If Mid(myXstrTmpPath, i, 1) = "\" Then n = n + 1
    Next i
    PfnclonCountFolderLayer = n
End Function

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｆ_フォルダ名を自動生成する
Private Function PfncstrFolderAutoNaming(ByVal myXstrDirPath As String) As String
    PfncstrFolderAutoNaming = Empty
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisDirExistFlag As Boolean
    myXbisDirExistFlag = myXobjFSO.FolderExists(myXstrDirPath)
    Do While myXbisDirExistFlag = True
        myXstrDirPath = myXstrDirPath & "_" & Format(Now, "YYYYMMDDHHmmSS")
        myXbisDirExistFlag = myXobjFSO.FolderExists(myXstrDirPath)
    Loop
    PfncstrFolderAutoNaming = myXstrDirPath
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_指定フォルダのフォルダ名情報を取得する
Private Sub PfixGetFolderNameInformation( _
            myXstrPrntPath As String, myXstrDirName As String, _
            ByVal myXstrDirPath As String)
    myXstrPrntPath = Empty: myXstrDirName = Empty
    If myXstrDirPath = "" Then Exit Sub
  Dim i As Long, myXlonInStrRev As Long: myXlonInStrRev = 0
    For i = Len(myXstrDirPath) To 1 Step -1
        If Mid(myXstrDirPath, i, 1) = "\" Then
            myXlonInStrRev = i: Exit For
        End If
    Next i
    If myXlonInStrRev <= 0 Or myXlonInStrRev >= Len(myXstrDirPath) Then Exit Sub
    myXstrPrntPath = Left(myXstrDirPath, myXlonInStrRev - 1)
    myXstrDirName = Mid(myXstrDirPath, myXlonInStrRev + 1)
End Sub

 '定型Ｐ_フォルダ階層をチェックしてフォルダを新規作成する
Private Sub PfixCreateFolderWithHierarchyCheck(myXbisExitFlag As Boolean, _
            ByVal myXstrBaseDirPath As String, ByVal myXstrRltvDirPath As String)
    myXbisExitFlag = False
    If myXstrRltvDirPath = "" Then GoTo ExitPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
'//基点フォルダが存在しなければ中止
    If myXobjFSO.FolderExists(myXstrBaseDirPath) = False Then GoTo ExitPath
'//新規フォルダと同一名フォルダが存在すれば中止
  Dim myXstrTmpPath As String
    myXstrTmpPath = myXstrBaseDirPath & "\" & myXstrRltvDirPath
    If myXobjFSO.FolderExists(myXstrTmpPath) = True Then Exit Sub
'//フォルダパスをフォルダ名に分割
  Dim myZvarSplitVrbl As Variant
    myZvarSplitVrbl = Split(myXstrRltvDirPath, "\")
'//フォルダの有無を階層で順次チェックして、フォルダ無しの場合フォルダを新規作成
    myXstrTmpPath = myXstrBaseDirPath
  Dim i As Long
    For i = LBound(myZvarSplitVrbl) To UBound(myZvarSplitVrbl)
        myXstrTmpPath = myXstrTmpPath & "\" & myZvarSplitVrbl(i)
        With myXobjFSO
            If .FolderExists(myXstrTmpPath) = True Then GoTo NextPath
            Call .CreateFolder(myXstrTmpPath)
        End With
NextPath:
    Next i
    Set myXobjFSO = Nothing: myZvarSplitVrbl = Empty
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｐ_フォルダを新規作成する
Private Sub PfixCreateFolder(myXbisExitFlag As Boolean, _
            ByVal myXstrDirPath As String)
    myXbisExitFlag = False
    If myXstrDirPath = "" Then GoTo ExitPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FolderExists(myXstrDirPath) = True Then GoTo ExitPath
        On Error GoTo ExitPath
        Call .CreateFolder(myXstrDirPath)
        On Error GoTo 0
    End With
    Set myXobjFSO = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｐ_Textファイルを新規作成する
Private Sub PfixCreateTextFile(myXbisExitFlag As Boolean, _
            ByVal myXstrFilePath As String)
    myXbisExitFlag = False
    If myXstrFilePath = "" Then GoTo ExitPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FileExists(myXstrFilePath) = True Then GoTo ExitPath
        On Error GoTo ExitPath
        Call .CreateTextFile(myXstrFilePath)
        On Error GoTo 0
    End With
    Set myXobjFSO = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '定型Ｐ_指定ファイルのファイル名情報を取得する(FileSystemObject使用)
Private Sub PfixGetFileNameInformationByFSO( _
            myXstrPrntPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String, _
            ByVal myXstrFilePath As String)
    myXstrPrntPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    If myXstrFilePath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        myXstrPrntPath = .GetParentFolderName(myXstrFilePath)   '親フォルダパス
        myXstrFileName = .GetFileName(myXstrFilePath)           'ファイル名
        myXstrBaseName = .GetBaseName(myXstrFilePath)           'ファイルベース名
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         'ファイル拡張子
    End With
    Set myXobjFSO = Nothing
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_テキストファイルを新規作成する
'Private Sub instCFileCreate()
'  Dim myXbisCmpltFlag As Boolean
'  Dim Lc As Long
'  Dim myXlonExeFileCnt As Long, _
'        myZstrExeFilePath() As String, myZstrExeFileName() As String
'    'myZstrExeFilePath(k) : 実行ファイルパス
'    'myZstrExeFileName(k) : 実行ファイル名
''//クラス内変数への入力
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'  Dim i As Long
'  Dim myXlonCreateFileCnt As Long, myZstrCreateFilePath() As String
'    'myZstrCreateFilePath(k) : 作成ファイルパス
'    myXlonCreateFileCnt = 2
'    i = myXlonCreateFileCnt + L - 1: ReDim myZstrCreateFilePath(i) As String
'    myZstrCreateFilePath(1) = ThisWorkbook.Path & "\test1" & ".txt"
'    myZstrCreateFilePath(2) = ThisWorkbook.Path & "\test2" & ".txt"
''  Dim myXlonCreateFileCnt As Long, myZstrCreateFilePath() As String
''    'myZstrCreateFilePath(k) : 作成ファイルパス
''    myXlonCreateFileCnt = XarbLong
''    i = myXlonCreateFileCnt + L - 1: ReDim myZstrCreateFilePath(i) As String
''    myZstrCreateFilePath(1) = XarbString
''//CFileCreateクラスモジュールのインスタンス生成
'  Dim myXinsFileCreate As CFileCreate: Set myXinsFileCreate = New CFileCreate
'    With myXinsFileCreate
'    '//クラス内変数への入力
'        .letAutoNaming = True
'        If myXlonCreateFileCnt <= 0 Then GoTo JumpPath
'        .letCreateFileCnt = myXlonCreateFileCnt
'        .letCreateFilePathAry() = myZstrCreateFilePath()
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXbisCmpltFlag = .getCmpltFlag
'        myXlonExeFileCnt = .getExeFileCnt
'        If myXlonExeFileCnt <= 0 Then GoTo JumpPath
'        i = myXlonExeFileCnt + L - 1
'        ReDim myZstrExeFilePath(i) As String
'        ReDim myZstrExeFileName(i) As String
'        Lc = .getOptnBase
'        For i = 1 To myXlonExeFileCnt
'            myZstrExeFilePath(i + L - 1) = .getExeFilePathAry(i + Lc - 1)
'            myZstrExeFileName(i + L - 1) = .getExeFileNameAry(i + Lc - 1)
'        Next i
'    End With
'    Set myXinsFileCreate = Nothing
'    Call variablesOfCFileCreate(myXlonExeFileCnt, myZstrExeFilePath)     'Debug.Print
'JumpPath:
'End Sub
'Private Sub variablesOfCFileCreate( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CFileCreateクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
