VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSubFileLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncobjGetFolder
'Includes PincGetNotHiddenSubFileArray
'Includes PfncbisCheckFileAttribute
'Includes PincPickUpExtensionMatchFileArray
'Includes PfncbisCheckFileExtension
'Includes PincSortFileArray
'Includes PfixGetFolderFileStringInformationAndObject2DArray
'Includes PfixQuickSortOfStringObject2DArray
'Includes PincSortFileArrayDateLastModified
'Includes PfixGetFolderFileDateInformationAndObject2DArray
'Includes PfixQuickSortOfNumericObject2DArray
'Includes PfixGetFolderFileStringInformationFor1DArray

Option Explicit
Option Base 1

'◆ClassProc名_指定ディレクトリ内のサブファイル一覧を取得する
'Rev.007

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSubFileLst"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
            myZstrFileName() As String, myZstrFilePath() As String
  
'//PropertySetLet_入力制御信号
  Private myXlonFileSortOptn As Long, myXbisNotOutFileInfo As Boolean
  
'//PropertySetLet_入力データ
  Private myXstrDirPath As String, myXstrExtsn As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjDir As Object
  Private myZobjNotHdnFile() As Object, myZobjExtMtchFile() As Object

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjDir = Nothing
    Erase myZobjNotHdnFile: Erase myZobjExtMtchFile
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getFileCnt() As Long
    getFileCnt = myXlonFileCnt
End Property

Public Property Get getFileAry(ByVal k As Long) As Object
    Set getFileAry = myZobjFile(k)
'myZobjFile(k) : ファイルオブジェクト
End Property

Public Property Get getFileNameAry(ByVal k As Long) As String
    getFileNameAry = myZstrFileName(k)
'myZstrFileName(k) : ファイル名
End Property

Public Property Get getFilePathAry(ByVal k As Long) As String
    getFilePathAry = myZstrFilePath(k)
'myZstrFilePath(k) : ファイルパス
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letFileSortOptn(ByVal RHS As Long)
    myXlonFileSortOptn = RHS
'myXlonFileSortOptn = 1 : ソートしない
'myXlonFileSortOptn = 2 : ファイル名順にソートする
'myXlonFileSortOptn = 3 : 更新日時順にソートする
End Property

Public Property Let letNotOutFileInfo(ByVal RHS As Boolean)
    myXbisNotOutFileInfo = RHS
'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
End Property

Public Property Let letDirPath(ByVal RHS As String)
    myXstrDirPath = RHS
End Property

Public Property Let letExtsn(ByVal RHS As String)
    myXstrExtsn = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXlonFileSortOptn < 1 Or myXlonFileSortOptn > 3 Then
        myXlonFileSortOptn = 2
    End If
    
    If myXstrDirPath = "" Then GoTo ExitPath
    If InStr(myXstrExtsn, ".") <> 0 Then _
        myXstrExtsn = Mid(myXstrExtsn, InStr(myXstrExtsn, ".") + 1)
    
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
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//S:指定ディレクトリ内の指定拡張子のサブファイル一覧を取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:ファイル一覧のソート処理
    Call prsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:ファイル情報取得
    Call prsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'SnsP_指定ディレクトリ内の指定拡張子のサブファイル一覧を取得
Private Sub snsProc()
    myXbisExitFlag = False
    
'//指定フォルダのオブジェクトを取得
    Set myXobjDir = PfncobjGetFolder(myXstrDirPath)
    If myXobjDir Is Nothing Then GoTo ExitPath
    
'//指定ディレクトリ内の隠しファイル以外のサブファイルオブジェクト一覧を取得
    Call PincGetNotHiddenSubFileArray(myXlonFileCnt, myZobjNotHdnFile, myXobjDir)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
'//ファイル一覧から指定拡張子のファイルを抽出
    Call PincPickUpExtensionMatchFileArray( _
            myXlonFileCnt, myZobjExtMtchFile, _
            myZobjNotHdnFile, myXstrExtsn)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_ファイル一覧のソート処理
Private Sub prsProc1()
    myXbisExitFlag = False
    
  Dim i As Long
    If myXlonFileSortOptn = 1 Then
    '//ファイル一覧をソート処理しない
        ReDim myZobjFile(myXlonFileCnt) As Object
        For i = LBound(myZobjExtMtchFile) To UBound(myZobjExtMtchFile)
            Set myZobjFile(i) = myZobjExtMtchFile(i)
        Next i
        
    ElseIf myXlonFileSortOptn = 2 Then
    '//ファイル一覧をファイル名順にソート
        Call PincSortFileArray(myXlonFileCnt, myZobjFile, myZobjExtMtchFile)
        If myXlonFileCnt <= 0 Then GoTo ExitPath
        
    ElseIf myXlonFileSortOptn = 3 Then
    '//ファイル一覧を更新日時順にソート
        Call PincSortFileArrayDateLastModified(myXlonFileCnt, myZobjFile, myZobjExtMtchFile)
        If myXlonFileCnt <= 0 Then GoTo ExitPath
        
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_ファイル情報取得
Private Sub prsProc2()
    myXbisExitFlag = False
    
    If myXbisNotOutFileInfo = True Then Exit Sub
    
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
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

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

 '定型Ｐ_指定ディレクトリ内の隠しファイル以外のサブファイルオブジェクト一覧を取得する
Private Sub PincGetNotHiddenSubFileArray( _
            myXlonFileCnt As Long, myZobjNotHdnFile() As Object, _
            ByVal myXobjDir As Object, _
            Optional ByVal coXlonAtrbt As Long = 2)
'Includes PfncbisCheckFileAttribute
'myZobjNotHdnFile(i) : サブファイル
'coXlonAtrbt = 2 : 隠しファイル
    myXlonFileCnt = Empty: Erase myZobjNotHdnFile
    If myXobjDir Is Nothing Then Exit Sub
    If myXobjDir.Files.Count <= 0 Then Exit Sub
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXobjSubFile As Object, myXbisAtrChck As Boolean, n As Long: n = L - 1
    For Each myXobjSubFile In myXobjDir.Files
        myXbisAtrChck = PfncbisCheckFileAttribute(myXobjSubFile, coXlonAtrbt, True)
        If myXbisAtrChck = True Then GoTo NextPath
        n = n + 1: ReDim Preserve myZobjNotHdnFile(n) As Object
        Set myZobjNotHdnFile(n) = myXobjSubFile
NextPath:
    Next myXobjSubFile
    myXlonFileCnt = n - L + 1
    Set myXobjSubFile = Nothing
End Sub

 '定型Ｆ_指定ファイルが指定属性であることを確認する(WFDec2Bin使用)
Private Function PfncbisCheckFileAttribute( _
            ByVal myXobjFile As Object, ByVal myXlonAtrbt As Long, _
            Optional ByVal coXbisInStrOptn As Boolean = False) As Boolean
'myXlonAtrbt = 0   : 0000 0000 : 標準ファイル
'myXlonAtrbt = 1   : 0000 0001 : 読み取り専用ファイル
'myXlonAtrbt = 2   : 0000 0010 : 隠しファイル
'myXlonAtrbt = 4   : 0000 0100 : システムファイル
'myXlonAtrbt = 8   : 0000 1000 : ディスクドライブボリュームラベル(取得のみ可能)
'myXlonAtrbt = 16  : 0001 0000 : フォルダまたはディレクトリ(取得のみ可能)
'myXlonAtrbt = 32  : 0010 0000 : アーカイブファイル
'myXlonAtrbt = 64  : 0100 0000 : リンクまたはショートカット(取得のみ可能)
'myXlonAtrbt = 128 : 1000 0000 : 圧縮ファイル(取得のみ可能)
'※複数の属性を持つ場合には数値を足し合わせる
'coXbisInStrOptn = False : 値が一致すれば、myXbisCheck=True
'coXbisInStrOptn = True  : 値が含まれれば、myXbisCheck=True
    PfncbisCheckFileAttribute = False
    If myXobjFile Is Nothing Then Exit Function
    If myXlonAtrbt < 0 Then Exit Function
  Dim myXlonFileAtrbt As Long, myXlonFileAtrbtBin As String
    With myXobjFile
        myXlonFileAtrbt = .Attributes
        myXlonFileAtrbtBin = Application.WorksheetFunction.Dec2Bin(myXlonFileAtrbt, 8)
    End With
  Dim myXstrDsgntAtrbtBin As String
    Select Case coXbisInStrOptn
        Case True
            myXstrDsgntAtrbtBin = Application.WorksheetFunction.Dec2Bin(myXlonAtrbt, 8)
          Dim i As Long
            For i = Len(myXstrDsgntAtrbtBin) To 1 Step -1
                If Mid(myXstrDsgntAtrbtBin, i, 1) = 1 Then _
                    If Mid(myXlonFileAtrbtBin, i, 1) <> 1 Then Exit Function
            Next i
        Case Else
            If myXlonFileAtrbt <> myXlonAtrbt Then Exit Function
    End Select
    PfncbisCheckFileAttribute = True
End Function

 '定型Ｐ_ファイル一覧から指定拡張子と一致するファイルを抽出する
Private Sub PincPickUpExtensionMatchFileArray( _
            myXlonExtMtchFileCnt As Long, myZobjExtMtchFile() As Object, _
            ByRef myZobjOrgFile() As Object, ByVal myXstrExtsn As String)
'Includes PfncbisCheckFileExtension
'myZobjExtMtchFile(i) : 抽出ファイル
'myZobjOrgFile(i) : 元ファイル
    myXlonExtMtchFileCnt = Empty: Erase myZobjExtMtchFile
  Dim myXobjTmp As Object, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZobjOrgFile): Set myXobjTmp = myZobjOrgFile(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarFile As Variant, myXbisExtChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarFile In myZobjOrgFile
      Dim myXstrFilePath As String: myXstrFilePath = myXvarFile.Path
        myXbisExtChck = PfncbisCheckFileExtension(myXstrFilePath, myXstrExtsn)
        If myXbisExtChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZobjExtMtchFile(n) As Object
        Set myZobjExtMtchFile(n) = myXvarFile
NextPath:
    Next
    myXlonExtMtchFileCnt = n - Lo + 1
    myXvarFile = Empty
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

 '定型Ｐ_ファイル一覧をファイル名順にソートする
Private Sub PincSortFileArray( _
            myXlonFileCnt As Long, myZobjFile() As Object, _
            ByRef myZobjOrgFile() As Object, _
            Optional ByVal coXlonOrder As Long = 1, _
            Optional ByVal coXlonMode As Long = 1)
'Includes PfixGetFolderFileStringInformationAndObject2DArray
'Includes PfixQuickSortOfStringObject2DArray
'myZobjOrgFile(i) : 元ファイル
'coXlonOrder = 1  : 昇順 (Ascending Order)
'coXlonOrder = -1 : 降順 (Descending Order)
'coXlonMode = vbBinaryCompare  = 0 : バイナリモードの比較
'coXlonMode = vbTextCompare    = 1 : テキストモードの比較
    myXlonFileCnt = Empty: Erase myZobjFile
'//文字列情報一覧を取得
  Dim myXlonInfoCnt As Long, myZvarStrAndObj() As Variant
    Call PfixGetFolderFileStringInformationAndObject2DArray( _
            myXlonInfoCnt, myZvarStrAndObj, _
            myZobjOrgFile, 1)
    If myXlonInfoCnt <= 0 Then Exit Sub
    myXlonFileCnt = myXlonInfoCnt
'//文字列によるクイックソート
  Dim L As Long, U As Long
    L = LBound(myZvarStrAndObj, 1)
    U = UBound(myZvarStrAndObj, 1)
    Call PfixQuickSortOfStringObject2DArray(myZvarStrAndObj, L, U, coXlonOrder, coXlonMode)
'//オブジェクトを配列変数に格納
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long: i = myXlonFileCnt + Lo - 1
    ReDim myZobjFile(i) As Object
    For i = LBound(myZvarStrAndObj, 1) To UBound(myZvarStrAndObj, 1)
        Set myZobjFile(i + Lo - L) = myZvarStrAndObj(i, L + 1)
    Next i
    Erase myZvarStrAndObj
End Sub

 '定型Ｐ_フォルダファイルオブジェクト一覧の文字列情報とオブジェクト併記の配列を取得する
Private Sub PfixGetFolderFileStringInformationAndObject2DArray( _
            myXlonInfoCnt As Long, myZvarStrInfoAndObj() As Variant, _
            ByRef myZobjFldrFile() As Object, _
            Optional ByVal coXlonStrOptn As Long = 1)
'myZvarStrInfoAndObj(i, 1) : 文字列情報
'myZvarStrInfoAndObj(i, 2) : 元フォルダor元ファイル
'myZobjFldrFile(i) : 元フォルダor元ファイル
'coXlonStrOptn = 1  : 名前 (Name)
'coXlonStrOptn = 2  : パス (Path)
'coXlonStrOptn = 3  : 親フォルダ (ParentFolder)
'coXlonStrOptn = 4  : 属性 (Attributes)
'coXlonStrOptn = 5  : 種類 (Type)
    myXlonInfoCnt = Empty: Erase myZvarStrInfoAndObj
  Dim myXobjTmp As Object, L As Long
    On Error GoTo ExitPath
    L = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(L)
    On Error GoTo 0
    myXlonInfoCnt = UBound(myZobjFldrFile) - LBound(myZobjFldrFile) + 1
  Dim i As Long: i = myXlonInfoCnt - (1 - L)
    ReDim myZvarStrInfoAndObj(i, L + 1) As Variant
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        If myZobjFldrFile(i) Is Nothing Then GoTo NextPath
        Select Case coXlonStrOptn
            Case 1: myZvarStrInfoAndObj(i, L) = myZobjFldrFile(i).Name
            Case 2: myZvarStrInfoAndObj(i, L) = myZobjFldrFile(i).Path
            Case 3: myZvarStrInfoAndObj(i, L) = myZobjFldrFile(i).ParentFolder
            Case 4: myZvarStrInfoAndObj(i, L) = myZobjFldrFile(i).Attributes
            Case 5: myZvarStrInfoAndObj(i, L) = myZobjFldrFile(i).Type
        End Select
        Set myZvarStrInfoAndObj(i, L + 1) = myZobjFldrFile(i)
NextPath:
    Next i
ExitPath:
End Sub

 '定型Ｐ_オブジェクト配列の文字列によるクイックソート
Private Sub PfixQuickSortOfStringObject2DArray( _
            myZvarStrAndObj As Variant, _
            ByVal myXlonLbnd As Long, ByVal myXlonUbnd As Long, _
            Optional ByVal coXlonOrder As Long = 1, _
            Optional ByVal coXlonMode As Long = 1)
'myZvarStrAndObj(i, 1) : 文字列情報
'myZvarStrAndObj(i, 2) : オブジェクト
'coXlonOrder = 1  : 昇順 (Ascending Order)
'coXlonOrder = -1 : 降順 (Descending Order)
'coXlonMode = vbBinaryCompare  = 0 : バイナリモードの比較
'coXlonMode = vbTextCompare    = 1 : テキストモードの比較
    If myXlonUbnd <= myXlonLbnd Then Exit Sub
  Dim i As Long, j As Long, L As Long
    On Error GoTo ExitPath
    L = LBound(myZvarStrAndObj, 2)
    On Error GoTo 0
  Dim myXstrMed As String, myXstrTmp1 As String, myXobjTmp2 As Object
    myXstrMed = myZvarStrAndObj(Int((myXlonLbnd + myXlonUbnd) / 2), L)
    i = myXlonLbnd: j = myXlonUbnd
    Do
        Do While StrComp(myZvarStrAndObj(i, L), myXstrMed, coXlonMode) _
                    = -1 * coXlonOrder
            i = i + 1
        Loop
        Do While StrComp(myZvarStrAndObj(j, L), myXstrMed, coXlonMode) _
                    = coXlonOrder
            j = j - 1
        Loop
        If i >= j Then Exit Do
        myXstrTmp1 = CStr(myZvarStrAndObj(i, L + 0))
        Set myXobjTmp2 = myZvarStrAndObj(i, L + 1)
        myZvarStrAndObj(i, L + 0) = myZvarStrAndObj(j, L + 0)
        Set myZvarStrAndObj(i, L + 1) = myZvarStrAndObj(j, L + 1)
        myZvarStrAndObj(j, L + 0) = myXstrTmp1
        Set myZvarStrAndObj(j, L + 1) = myXobjTmp2
        i = i + 1: j = j - 1
    Loop
    If myXlonLbnd < i - 1 Then _
        Call PfixQuickSortOfStringObject2DArray( _
                myZvarStrAndObj, myXlonLbnd, i - 1, coXlonOrder, coXlonMode)
    If myXlonUbnd > j + 1 Then _
        Call PfixQuickSortOfStringObject2DArray( _
                myZvarStrAndObj, j + 1, myXlonUbnd, coXlonOrder, coXlonMode)
ExitPath:
End Sub

 '定型Ｐ_ファイル一覧を更新日時順にソートする
Private Sub PincSortFileArrayDateLastModified( _
            myXlonFileCnt As Long, myZobjFile() As Object, _
            ByRef myZobjOrgFile() As Object, _
            Optional ByVal coXlonOrder As Long = -1)
'Includes PfixGetFolderFileDateInformationAndObject2DArray
'Includes PfixQuickSortOfNumericObject2DArray
'myZobjOrgFile(i) : 元ファイル
'coXlonOrder = 1    : 昇順 (Ascending Order)
'coXlonOrder = -1   : 降順 (Descending Order)
    myXlonFileCnt = Empty: Erase myZobjFile
'//更新日時情報一覧を取得
  Dim myXlonInfoCnt As Long, myZvarDatInfoAndObj() As Variant
    Call PfixGetFolderFileDateInformationAndObject2DArray( _
            myXlonInfoCnt, myZvarDatInfoAndObj, _
            myZobjOrgFile, 3)
    If myXlonInfoCnt <= 0 Then Exit Sub
    myXlonFileCnt = myXlonInfoCnt
'//数値によるクイックソート
  Dim L As Long, U As Long
    L = LBound(myZvarDatInfoAndObj, 1)
    U = UBound(myZvarDatInfoAndObj, 1)
    Call PfixQuickSortOfNumericObject2DArray(myZvarDatInfoAndObj, L, U, coXlonOrder)
'//オブジェクトを配列変数に格納
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long: i = myXlonFileCnt + Lo - 1
    ReDim myZobjFile(i) As Object
    For i = LBound(myZvarDatInfoAndObj, 1) To UBound(myZvarDatInfoAndObj, 1)
        Set myZobjFile(i + Lo - L) = myZvarDatInfoAndObj(i, L + 1)
    Next i
    Erase myZvarDatInfoAndObj
End Sub

 '定型Ｐ_フォルダファイルオブジェクト一覧の日付情報とオブジェクト併記の配列を取得する
Private Sub PfixGetFolderFileDateInformationAndObject2DArray( _
            myXlonInfoCnt As Long, myZvarDatInfoAndObj() As Variant, _
            ByRef myZobjFldrFile() As Object, _
            Optional ByVal coXlonDateOptn As Long = 3)
'myZvarDatInfoAndObj(i, 1) : 日付情報
'myZvarDatInfoAndObj(i, 2) : 元フォルダor元ファイル
'myZobjFldrFile(i) : 元フォルダor元ファイル
'coXlonDateOptn = 1  : 作成日時 (DateCreated)
'coXlonDateOptn = 2  : 最終アクセス日時 (DateLastAccessed)
'coXlonDateOptn = 3  : 最終更新日時 (DateLastModified)
    myXlonInfoCnt = Empty: Erase myZvarDatInfoAndObj
  Dim myXobjTmp As Object, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(Li)
    On Error GoTo 0
  Dim myZdatBuf() As Date, myZobjBuf() As Object
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXdatTmp As Date, i As Long, n As Long: n = Lo - 1
    On Error GoTo NextPath
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        myXdatTmp = Empty
        Select Case coXlonDateOptn
            Case 1: myXdatTmp = myZobjFldrFile(i).DateCreated       '作成日時
            Case 2: myXdatTmp = myZobjFldrFile(i).DateLastAccessed  'アクセス日時
            Case 3: myXdatTmp = myZobjFldrFile(i).DateLastModified  '更新日時
        End Select
        n = n + 1
        ReDim Preserve myZdatBuf(n) As Date
        ReDim Preserve myZobjBuf(n) As Object
        myZdatBuf(n) = myXdatTmp
        Set myZobjBuf(n) = myZobjFldrFile(i)
NextPath:
    Next i
    On Error GoTo 0
    myXlonInfoCnt = n - Lo + 1
    ReDim myZvarDatInfoAndObj(n, Lo + 1) As Variant
    For i = LBound(myZdatBuf) To UBound(myZdatBuf)
        myZvarDatInfoAndObj(i, Lo + 0) = myZdatBuf(i)
        Set myZvarDatInfoAndObj(i, Lo + 1) = myZobjBuf(i)
    Next i
    Erase myZdatBuf: Erase myZobjBuf
ExitPath:
End Sub

 '定型Ｐ_オブジェクト配列の数値によるクイックソート
Private Sub PfixQuickSortOfNumericObject2DArray( _
            myZvarNumAndObj As Variant, _
            ByVal myXlonLbnd As Long, ByVal myXlonUbnd As Long, _
            Optional ByVal coXlonOrder As Long = 1)
'myZvarNumAndObj(i, 1) : 数値情報
'myZvarNumAndObj(i, 2) : オブジェクト
'coXlonOrder = 1    : 昇順 (Ascending Order)
'coXlonOrder = -1   : 降順 (Descending Order)
    If myXlonUbnd <= myXlonLbnd Then Exit Sub
  Dim i As Long, j As Long, L As Long
    On Error GoTo ExitPath
    L = LBound(myZvarNumAndObj, 2)
    On Error GoTo 0
  Dim myXdouMed As Double, myXdouTmp1 As Double, myXobjTmp2 As Object
    myXdouMed = myZvarNumAndObj(Int((myXlonLbnd + myXlonUbnd) / 2), L)
    i = myXlonLbnd: j = myXlonUbnd
    Do
        Do While coXlonOrder * (myZvarNumAndObj(i, L) - myXdouMed) < 0
            i = i + 1
        Loop
        Do While coXlonOrder * (myZvarNumAndObj(j, L) - myXdouMed) > 0
            j = j - 1
        Loop
        If i >= j Then Exit Do
        myXdouTmp1 = CDbl(myZvarNumAndObj(i, L + 0))
        Set myXobjTmp2 = myZvarNumAndObj(i, L + 1)
        myZvarNumAndObj(i, L + 0) = myZvarNumAndObj(j, L + 0)
        Set myZvarNumAndObj(i, L + 1) = myZvarNumAndObj(j, L + 1)
        myZvarNumAndObj(j, L + 0) = myXdouTmp1
        Set myZvarNumAndObj(j, L + 1) = myXobjTmp2
        i = i + 1: j = j - 1
        Set myXobjTmp2 = Nothing
    Loop
    If myXlonLbnd < i - 1 Then _
        Call PfixQuickSortOfNumericObject2DArray( _
                myZvarNumAndObj, myXlonLbnd, i - 1, coXlonOrder)
    If myXlonUbnd > j + 1 Then _
        Call PfixQuickSortOfNumericObject2DArray( _
                myZvarNumAndObj, j + 1, myXlonUbnd, coXlonOrder)
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
  Dim myXobjTmp As Object, L As Long
    On Error GoTo ExitPath
    L = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(L)
    On Error GoTo 0
    myXlonInfoCnt = UBound(myZobjFldrFile) - LBound(myZobjFldrFile) + 1
  Dim i As Long: i = myXlonInfoCnt - (1 - L)
    ReDim myZstrInfo(i) As String
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        If myZobjFldrFile(i) Is Nothing Then GoTo NextPath
        Select Case coXlonStrOptn
            Case 1: myZstrInfo(i) = myZobjFldrFile(i).Name
            Case 2: myZstrInfo(i) = myZobjFldrFile(i).Path
            Case 3: myZstrInfo(i) = myZobjFldrFile(i).ParentFolder
            Case 4: myZstrInfo(i) = myZobjFldrFile(i).Attributes
            Case 5: myZstrInfo(i) = myZobjFldrFile(i).Type
        End Select
NextPath:
    Next i
ExitPath:
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_指定ディレクトリ内のサブファイル一覧を取得する
'Private Sub instCSubFileLst()
'  Dim Lc As Long
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String
'    'myZobjFile(k) : ファイルオブジェクト
'    'myZstrFileName(k) : ファイル名
'    'myZstrFilePath(k) : ファイルパス
''//クラス内変数への入力
'  Dim myXlonFileSortOptn As Long
'    'myXlonFileSortOptn = 1 : ソートしない
'    'myXlonFileSortOptn = 2 : ファイル名順にソートする
'    'myXlonFileSortOptn = 3 : 更新日時順にソートする
'    myXlonFileSortOptn = 1
'  Dim myXbisNotOutFileInfo As Boolean
'    'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
'    'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
'    myXbisNotOutFileInfo = False
'  Dim myXstrDirPath As String
'    myXstrDirPath = ActiveWorkbook.Path
'  Dim myXstrExtsn As String: myXstrExtsn = ""
''//CSubFileLstクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim k As Long
'  Dim myXinsSubFileLst As CSubFileLst: Set myXinsSubFileLst = New CSubFileLst
'    With myXinsSubFileLst
'    '//クラス内変数への入力
'        .letFileSortOptn = myXlonFileSortOptn
'        .letNotOutFileInfo = myXbisNotOutFileInfo
'        .letDirPath = myXstrDirPath
'        .letExtsn = myXstrExtsn
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonFileCnt = .getFileCnt
'        If myXlonFileCnt <= 0 Then GoTo JumpPath
'        k = myXlonFileCnt + Lo - 1
'        ReDim myZobjFile(k) As Object
'        ReDim myZstrFileName(k) As String
'        ReDim myZstrFilePath(k) As String
'        Lc = .getOptnBase
'        For k = 1 To myXlonFileCnt
'            Set myZobjFile(k + Lo - 1) = .getFileAry(k + Lc - 1)
'            myZstrFileName(k + Lo - 1) = .getFileNameAry(k + Lc - 1)
'            myZstrFilePath(k + Lo - 1) = .getFilePathAry(k + Lc - 1)
'        Next k
'    End With
'JumpPath:
'    Set myXinsSubFileLst = Nothing
'    Call variablesOfCSubFileLst(myXlonFileCnt, myZstrFileName)   'Debug.Print
'End Sub
'Private Sub variablesOfCSubFileLst( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSubFileLstクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
