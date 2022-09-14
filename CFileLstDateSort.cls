VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileLstDateSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfixPickUpExistFileArray
'Includes PincSortFileArrayDateLastModified
'Includes PfixGetFolderFileDateInformationAndObject2DArray
'Includes PfixQuickSortOfNumericObject2DArray
'Includes PfixGetFolderFileStringInformationFor1DArray

Option Explicit
Option Base 1

'◆ClassProc名_ファイル一覧から更新日時順でファイル一覧を取得する
'Rev.003
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CFileLstDateSort"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
            myZstrFileName() As String, myZstrFilePath() As String
  
'//PropertySetLet_入力制御信号
  Private myXlonLastToOld As Long, myXbisNotOutFileInfo As Boolean
  
'//PropertySetLet_入力データ
  Private myZobjOrgFile() As Object, myZstrOrgFilePath() As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonOrgFileCnt As Long, _
            myZobjOrgFileInt() As Object, myZstrOrgFilePathInt() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgFileCnt = Empty: Erase myZobjOrgFileInt: Erase myZstrOrgFilePathInt
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getFileCnt() As Long
    getFileCnt = myXlonFileCnt
End Property

Public Property Get getFileNameAry(ByVal k As Long) As String
    getFileNameAry = myZstrFileName(k)
'myZstrFileName(k) : ファイル名
End Property

Public Property Get getFilePathAry(ByVal k As Long) As String
    getFilePathAry = myZstrFilePath(k)
'myZstrFilePath(k) : ファイルパス
End Property

Public Property Get getFileAry(ByVal k As Long) As Object
    Set getFileAry = myZobjFile(k)
'myZobjFile(k) : ファイルオブジェクト
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letLastToOld(ByVal RHS As Long)
    myXlonLastToOld = RHS
'myXlonLastToOld = -1 : 最新ファイルが先頭(降順)
'myXlonLastToOld = 1  : 古いファイルが先頭(昇順)
End Property

Public Property Let letNotOutFileInfo(ByVal RHS As Boolean)
    myXbisNotOutFileInfo = RHS
'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
End Property

Public Property Let letOrgFileAry(ByRef RHS() As Object)
'myZobjOrgFile(i) : 元ファイルオブジェクト配列
    myZobjOrgFile = RHS
End Property

Public Property Let letOrgFilePathAry(ByRef RHS() As String)
'myZstrOrgFilePath(i) : 元ファイルパス配列
    myZstrOrgFilePath = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables1()
    
  Dim myXobjTmp As Object, L As Long
    On Error GoTo ExitPath
    L = LBound(myZobjOrgFile): Set myXobjTmp = myZobjOrgFile(L)
    On Error GoTo 0
    
  Dim i As Long, n As Long: n = 0
    For i = LBound(myZobjOrgFile) To UBound(myZobjOrgFile)
        If myZobjOrgFile(i) Is Nothing Then GoTo NextPath
        n = n + 1: ReDim Preserve myZobjOrgFileInt(n) As Object
        Set myZobjOrgFileInt(n) = myZobjOrgFile(i)
NextPath:
    Next i
    myXlonOrgFileCnt = n
    
ExitPath:
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables2()
    myXbisExitFlag = False
    
    If myXlonOrgFileCnt > 0 Then Exit Sub
    
  Dim myXstrTmp As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgFilePath): myXstrTmp = myZstrOrgFilePath(L)
    On Error GoTo 0
    
'//ファイルパス一覧から存在するファイルを抽出
  Dim i As Long
    i = UBound(myZstrOrgFilePath) - LBound(myZstrOrgFilePath) + 1
    ReDim myZstrOrgFilePathInt(i) As String
    For i = LBound(myZstrOrgFilePath) To UBound(myZstrOrgFilePath)
        myZstrOrgFilePathInt(i + (1 - L)) = myZstrOrgFilePath(i)
    Next i
    Call PfixPickUpExistFileArray(myXlonOrgFileCnt, myZobjOrgFileInt, myZstrOrgFilePathInt)
    If myXlonOrgFileCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    Call ctrProc
End Sub

'PublicF_
Public Function fncobjLastModifiedFile() As Object
    Set fncobjLastModifiedFile = Nothing
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Function
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    Set fncobjLastModifiedFile = myZobjFile(L)
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call checkInputVariables1
    Call checkInputVariables2: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//定型Ｐ_ファイル一覧を更新日時順にソートする
    Call PincSortFileArrayDateLastModified( _
            myXlonFileCnt, myZobjFile, _
            myZobjOrgFileInt, myXlonLastToOld)
        If myXlonFileCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//ファイル情報取得
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'PrcsP_ファイル情報取得
Private Sub prsProc()
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

 '定型Ｐ_ファイルパス一覧から存在するファイルを抽出する
Private Sub PfixPickUpExistFileArray( _
            myXlonExistFileCnt As Long, myZobjExistFile() As Object, _
            ByRef myZstrOrgFilePath() As String)
'myZobjExistFile(i) : 抽出ファイル
'myZstrOrgFilePath(i) : 元ファイルパス
    myXlonExistFileCnt = Empty: Erase myZobjExistFile
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
        n = n + 1: ReDim Preserve myZobjExistFile(n) As Object
        Set myZobjExistFile(n) = myXobjFSO.GetFile(myXvarPath)
NextPath:
    Next myXvarPath
    myXlonExistFileCnt = n - Lo + 1
    Set myXobjFSO = Nothing
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

''◆ClassProc名_ファイル一覧から更新日時順でファイル一覧を取得する
'Private Sub instCFileLstDateSort()
'  Dim Lc As Long
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String
'    'myZobjFile(k) : ファイルオブジェクト
'    'myZstrFileName(k) : ファイル名
'    'myZstrFilePath(k) : ファイルパス
'  Dim myXobjLastModFile As Object
''//クラス内変数への入力
'  Dim myXlonLastToOld As Long
'    'myXlonLastToOld = -1 : 最新ファイルが先頭(降順)
'    'myXlonLastToOld = 1  : 古いファイルが先頭(昇順)
'    myXlonLastToOld = -1
'  Dim myXbisNotOutFileInfo As Boolean
'    'myXbisNotOutFileInfo = False : ファイルオブジェクとファイル情報を両方出力する
'    'myXbisNotOutFileInfo = True  : ファイルオブジェクトのみ出力してファイル情報は出力しない
'    myXbisNotOutFileInfo = False
'  Dim myZobjOrgFile() As Object, myZstrOrgFilePath() As String
'    'myZobjOrgFile(i) : 元ファイルオブジェクト配列
'    'myZstrOrgFilePath(i) : 元ファイルパス配列
'  Dim myXlonOrgFileCnt As Long
'    myXlonOrgFileCnt = 2
'    ReDim myZstrOrgFilePath(myXlonOrgFileCnt) As String
'    myZstrOrgFilePath(1) = "C:\Users\Hiroki\Documents\_VBA4XPC\11 プログラムデータベース\02_VBAモジュール\mdl01_マクロ作成サポート" & "\" & _
'                            "cls02a_VBE内の情報操作.xlsm"
'    myZstrOrgFilePath(2) = "C:\Users\Hiroki\Documents\_VBA4XPC\11 プログラムデータベース\02_VBAモジュール\mdl01_マクロ作成サポート" & "\" & _
'                            "cls02b_エクセルブック内の個別プロシージャを取得.xlsm"
''//CFileLstDateSortクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim k As Long
'  Dim myXinsFileLstDateSort As CFileLstDateSort
'    Set myXinsFileLstDateSort = New CFileLstDateSort
'    With myXinsFileLstDateSort
'    '//クラス内変数への入力
'        .letLastToOld = myXlonLastToOld
'        .letNotOutFileInfo = myXbisNotOutFileInfo
'        .letOrgFileAry = myZobjOrgFile
'        .letOrgFilePathAry = myZstrOrgFilePath
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
'        Set myXobjLastModFile = .fncobjLastModifiedFile
'    End With
'JumpPath:
'    Set myXinsFileLstDateSort = Nothing
'    Call variablesOfCFileLstDateSort(myXlonFileCnt, myZstrFileName)  'Debug.Print
'    Debug.Print "データ: " & myXobjLastModFile.Name
'End Sub
'Private Sub variablesOfCFileLstDateSort( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CFileLstDateSortクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
