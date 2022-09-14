VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSlctFilePathRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckArrayDimensionLength
'Includes PfncobjSetFileDialogForFile
'Includes PfncstrSetDefaultFolderPath
'Includes PfncstrSetDefaultFilePath
'Includes PfixSetFileDialogFixedFilter
'Includes PfixSetFileDialogVariableFilter
'Includes PfixSetFileDialogForFileInitial
'Includes PfixSelectFileByFileDialog
'Includes PincPickUpBaseNameMatchFilePathMatrixSomeConditions
'Includes PfncbisCheckFileBaseNameSomeConditions
'Includes PfixGetFileName2DArrayByFSO
'Includes PfixDataPutInClipboard

Option Explicit
Option Base 1

'◆ClassProc名_指定文字列を含むファイル名のファイルが選択されるまで繰返し選択してそのパスを取得する
'Rev.004

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CSlctFilePathRpt"
  
'//モジュール内定数
  Private Const coXlonFDType As Long = msoFileDialogFilePicker
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonFileCnt As Long, _
            myZstrFileName() As String, myZstrFilePath() As String
  
'//PropertySetLet_入力データ
  Private myXstrDfltFldrPath As String, myXstrDfltFilePath As String
  Private myXstrExtsn As String
  Private myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long
  Private myXlonIniView As Long, myXbisMultSlct As Boolean
  Private myZvarOdrTrgtWrdPos() As Variant
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjFileDialog As Object
  Private myXstrDfltPath As String
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
  Private myXlonOrdrCnt As Long, myXlonPckUpMax As Long, _
            myZlonBsMtchFileCnt() As Long, myZstrBsMtchFilePath() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjFileDialog = Nothing
    myXstrDfltPath = Empty
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
    myXlonOrdrCnt = Empty: myXlonPckUpMax = Empty
    Erase myZlonBsMtchFileCnt: Erase myZstrBsMtchFilePath
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getFileCnt() As Long
    getFileCnt = myXlonFileCnt
End Property

Public Property Get getFileNameAry(ByVal i As Long) As String
  Dim L As Long: L = LBound(myZstrFileName, 1)
    getFileNameAry = myZstrFileName(i, L)
'myZstrFileName(i, 1) : ファイル名
End Property

Public Property Get getFilePathAry(ByVal i As Long) As String
  Dim L As Long: L = LBound(myZstrFilePath, 1)
    getFilePathAry = myZstrFilePath(i, L)
'myZstrFilePath(i, 1) : ファイルパス
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty
    Erase myZstrFileName: Erase myZstrFilePath
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letDfltFldrPath(ByVal RHS As String)
    myXstrDfltFldrPath = RHS
'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'myXstrDfltFldrPath = "*" : デフォルトパスを指定
End Property

Public Property Let letDfltFilePath(ByVal RHS As String)
    myXstrDfltFilePath = RHS
'myXstrDfltFilePath = ""  : デフォルトパス指定無し
'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
'myXstrDfltFilePath = "*" : デフォルトパスを指定
End Property

Public Property Let letExtsn(ByVal RHS As String)
    myXstrExtsn = RHS
End Property

Public Property Let letAddFltr(ByRef RHS() As String)
    myZstrAddFltr = RHS
'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
End Property

Public Property Let letFltrClr(ByVal RHS As Boolean)
    myXbisFltrClr = RHS
'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
'myXbisFltrClr = True  : ファイルフィルタを初期化する
End Property

Public Property Let letFltrIndx(ByVal RHS As Long)
    myXlonFltrIndx = RHS
'myXlonFltrIndx = 1～ : ファイルフィルタの初期選択
End Property

Public Property Let letIniView(ByVal RHS As Long)
    myXlonIniView = RHS
'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
End Property

Public Property Let letMultSlct(ByVal RHS As Boolean)
    myXbisMultSlct = RHS
'myXbisMultSlct = False : 複数のファイルを選択不可能
'myXbisMultSlct = True  : 複数のファイルを選択可能
End Property

Public Property Let letOdrTrgtWrdPosAry(ByRef RHS() As Variant)
    myZvarOdrTrgtWrdPos = RHS
'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If myXlonIniView < 1 Or myXlonIniView > 9 Then myXlonIniView = msoFileDialogViewList
    
'//指定拡張子を確認
    If InStr(myXstrExtsn, ".") <> 0 Then
        myXstrExtsn = Mid(myXstrExtsn, InStr(myXstrExtsn, ".") + 1)
        myXstrExtsn = LCase(myXstrExtsn)
    End If
    
    If myXbisFltrClr = False Then
        If Left(myXstrExtsn, 2) = "xl" Then
            myXlonFltrIndx = 2
        ElseIf myXstrExtsn = "txt" Then
            myXlonFltrIndx = 3
        ElseIf myXstrExtsn = "csv" Then
            myXlonFltrIndx = 4
        ElseIf myXstrExtsn = "jpg" Then
            myXlonFltrIndx = 5
        ElseIf myXstrExtsn = "bmp" Then
            myXlonFltrIndx = 6
        ElseIf myXstrExtsn = "png" Then
            myXlonFltrIndx = 7
        Else
            myXlonFltrIndx = myXlonFltrIndx
        End If
    End If
    
'//入力配列変数の内容を確認
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXlonChckDmnsn As Long, myXlonChckLen() As Long
    myXlonChckDmnsn = 3
    ReDim myXlonChckLen(myXlonChckDmnsn + L - 1) As Long
    myXlonChckLen(L + 0) = 0
    myXlonChckLen(L + 1) = 0
    myXlonChckLen(L + 2) = 2
    If PfncbisCheckArrayDimensionLength( _
        myZvarOdrTrgtWrdPos, myXlonChckDmnsn, myXlonChckLen) = False Then GoTo ExitPath
    
  Dim n As Long
    n = UBound(myZvarOdrTrgtWrdPos, 1) - LBound(myZvarOdrTrgtWrdPos, 1) + 1
    If n > 1 Then myXbisMultSlct = True
    
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
    
'//C:FileDialogのデフォルト設定
    Call setProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
'//S:ダイアログボックスでファイルパス一覧を取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
'//クリップボードにデータを格納
  Dim L As Long
    If myXlonFileCnt = 1 Then
        L = LBound(myZstrFilePath, 1)
        Call PfixDataPutInClipboard(myZstrFilePath(L, L))
    End If
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"  'PassFlag
    
ExitPath:
    Call initializeModuleVariables
End Sub

'SetP_FileDialogのデフォルト設定
Private Sub setProc()
    myXbisExitFlag = False
    
'//FileDialogオブジェクトを設定
    Set myXobjFileDialog = PfncobjSetFileDialogForFile(coXlonFDType)
    If myXobjFileDialog Is Nothing Then GoTo ExitPath
    
'//デフォルトパスを設定
    myXstrDfltFldrPath = PfncstrSetDefaultFolderPath(myXstrDfltFldrPath)
    myXstrDfltFilePath = PfncstrSetDefaultFilePath(myXstrDfltFilePath)
    
    If myXstrDfltFldrPath <> "" And myXstrDfltFilePath <> "" Then
        myXstrDfltPath = myXstrDfltFilePath
    ElseIf myXstrDfltFldrPath = "" And myXstrDfltFilePath <> "" Then
        myXstrDfltPath = myXstrDfltFilePath
    ElseIf myXstrDfltFldrPath <> "" And myXstrDfltFilePath = "" Then
        myXstrDfltPath = myXstrDfltFldrPath
    Else
        myXstrDfltPath = ""
    End If
    
'//ファイル選択用のFileDialogオブジェクトのフィルタ固定値を設定
    Call PfixSetFileDialogFixedFilter(myXobjFileDialog, myXlonFltrIndx)
    
'//ファイル選択用のFileDialogオブジェクトのフィルタ可変値を設定
    Call PfixSetFileDialogVariableFilter( _
            myXobjFileDialog, _
            myZstrAddFltr, myXbisFltrClr, myXlonFltrIndx)
    
'//FileDialogオブジェクトの初期状態を設定
    Call PfixSetFileDialogForFileInitial( _
            myXobjFileDialog, _
            myXlonIniView, myXstrDfltPath, myXbisMultSlct)
    
    If myXobjFileDialog Is Nothing Then GoTo ExitPath
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_ダイアログボックスでファイルパス一覧を取得
Private Sub snsProc()
    myXbisExitFlag = False
  
  Dim i As Long, L As Long, U As Long
  Dim myXbisExitLpFlag As Boolean: myXbisExitLpFlag = False
    Do While myXbisExitLpFlag = False
        
    '//FileDialogオブジェクトを使用してファイルを選択
        Call PfixSelectFileByFileDialog(myXlonOrgFileCnt, myZstrOrgFilePath, myXobjFileDialog)
        If myXlonOrgFileCnt <= 0 Then Exit Do
        
    '//ファイル一覧から順番に複数のベース名条件を指定して一致するファイルパスを抽出
        Call PincPickUpBaseNameMatchFilePathMatrixSomeConditions( _
                myXlonOrdrCnt, myXlonPckUpMax, myZlonBsMtchFileCnt, myZstrBsMtchFilePath, _
                myZstrOrgFilePath, myZvarOdrTrgtWrdPos)
        If myXlonOrdrCnt <= 0 Or myXlonPckUpMax <= 0 Then GoTo NextPath
        
        L = LBound(myZstrBsMtchFilePath, 1)
        U = UBound(myZstrBsMtchFilePath, 1)
        For i = L To U
            If myZlonBsMtchFileCnt(i, L) <= 0 Then GoTo NextPath
        Next i
        
        ReDim myZstrFilePath(U, L) As String
        For i = L To U
            myZstrFilePath(i, L) = myZstrBsMtchFilePath(i, L)
        Next i
        myXbisExitLpFlag = True
        
NextPath:
    Loop
    If myXlonOrdrCnt <= 0 Or myXlonPckUpMax <= 0 Then GoTo ExitPath
    myXlonFileCnt = myXlonOrdrCnt
    
'//指定2次元ファイルパス一覧のファイル名一覧を取得
  Dim myXlonRows As Long, myXlonCols As Long
    Call PfixGetFileName2DArrayByFSO(myXlonRows, myXlonCols, myZstrFileName, myZstrFilePath)
    If myXlonRows <= 0 Or myXlonCols <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

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
        If myXlonChckAryLen(i) < 0 Then Exit Function
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

 '定型Ｆ_ファイル選択用のFileDialogオブジェクトを設定する
Private Function PfncobjSetFileDialogForFile( _
            Optional ByVal coXlonFDType As Long = msoFileDialogFilePicker) As Object
'coXlonFDType = msoFileDialogFilePicker : ファイルを選択
'coXlonFDType = msoFileDialogOpen       : ファイルを開く
'coXlonFDType = msoFileDialogSaveAs     : ファイルを保存
    Set PfncobjSetFileDialogForFile = Nothing
  Dim myXobjFileDialog As Object
    On Error GoTo ExitPath
    Set myXobjFileDialog = Application.FileDialog(coXlonFDType)
    myXobjFileDialog.Filters.Clear
    Select Case coXlonFDType
        Case msoFileDialogFilePicker
            myXobjFileDialog.Title = "ファイルの選択"
            myXobjFileDialog.ButtonName = "選択"
        Case msoFileDialogOpen
            myXobjFileDialog.Title = "ファイルを開く"
            myXobjFileDialog.ButtonName = "開く"
        Case msoFileDialogSaveAs
            myXobjFileDialog.Title = "ファイルを保存"
            myXobjFileDialog.ButtonName = "保存"
        Case Else: Exit Function
    End Select
    On Error GoTo 0
    Set PfncobjSetFileDialogForFile = myXobjFileDialog
    Set myXobjFileDialog = Nothing
ExitPath:
End Function

 '定型Ｆ_デフォルトフォルダパスを設定する
Private Function PfncstrSetDefaultFolderPath( _
            Optional ByVal coXstrDflt As String = "") As String
'coXstrDflt = ""  : デフォルトパス指定無し
'coXstrDflt = "C" : Cドライブをデフォルトパスに指定
'coXstrDflt = "1" : このブックの親フォルダをデフォルトパスに指定
'coXstrDflt = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'coXstrDflt = "*" : デフォルトパスを指定
    PfncstrSetDefaultFolderPath = Empty
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrDfltPath As String, myXbisDfltExist As Boolean
    With myXobjFSO
        Select Case coXstrDflt
            Case "C": myXstrDfltPath = "C:\"
            Case "1": myXstrDfltPath = .GetParentFolderName(ThisWorkbook.Path)
            Case "2": myXstrDfltPath = .GetParentFolderName(ActiveWorkbook.Path)
            Case Else: myXstrDfltPath = coXstrDflt
        End Select
        myXbisDfltExist = .FolderExists(myXstrDfltPath)
    End With
    If myXbisDfltExist = True Then PfncstrSetDefaultFolderPath = myXstrDfltPath
    Set myXobjFSO = Nothing
End Function

 '定型Ｆ_デフォルトファイルパスを設定する
Private Function PfncstrSetDefaultFilePath( _
            Optional ByVal coXstrDflt As String = "") As String
'coXstrDflt = ""  : デフォルトパス指定無し
'coXstrDflt = "1" : このブックをデフォルトパスに指定
'coXstrDflt = "2" : アクティブブックをデフォルトパスに指定
'coXstrDflt = "*" : デフォルトパスを指定
    PfncstrSetDefaultFilePath = Empty
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrDfltPath As String, myXbisDfltExist As Boolean
    Select Case coXstrDflt
        Case "1": myXstrDfltPath = ThisWorkbook.FullName
        Case "2": myXstrDfltPath = ActiveWorkbook.FullName
        Case Else: myXstrDfltPath = coXstrDflt
    End Select
    myXbisDfltExist = myXobjFSO.FileExists(myXstrDfltPath)
    If myXbisDfltExist = True Then PfncstrSetDefaultFilePath = myXstrDfltPath
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_ファイル選択用のFileDialogオブジェクトのフィルタ固定値を設定する
Private Sub PfixSetFileDialogFixedFilter( _
            myXobjFileDialog As Object, _
            Optional ByVal coXlonFltrIndx As Long = 1)
'coXlonFltrIndx = 1～ : ファイルフィルタの初期選択
  Const coXlonFDType As Long = msoFileDialogFilePicker  'ファイルを選択
    If myXobjFileDialog Is Nothing Then Exit Sub
  Dim i As Long: i = 0
    With myXobjFileDialog
        .Filters.Clear
        i = i + 1: Call .Filters.Add("全てのファイル", "*.*", i)
        i = i + 1: Call .Filters.Add("Excel File", "*.xl*", i)
        i = i + 1: Call .Filters.Add("Text File", "*.txt", i)
        i = i + 1: Call .Filters.Add("CSV File", "*.csv", i)
        i = i + 1: Call .Filters.Add("画像ファイル", "*.jpg", i)
        i = i + 1: Call .Filters.Add("画像ファイル", "*.bmp", i)
        i = i + 1: Call .Filters.Add("画像ファイル", "*.png", i)
        If coXlonFltrIndx < 1 Or coXlonFltrIndx > i Then coXlonFltrIndx = 1
        .FilterIndex = coXlonFltrIndx
    End With
End Sub

 '定型Ｐ_ファイル選択用のFileDialogオブジェクトのフィルタ可変値を設定する
Private Sub PfixSetFileDialogVariableFilter( _
            myXobjFileDialog As Object, _
            ByRef myZstrAddFltr() As String, _
            Optional ByVal coXbisFltrClr As Boolean = False, _
            Optional ByVal coXlonFltrIndx As Long = 1)
'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル説明)
'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ拡張子)
'coXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
'coXbisFltrClr = True  : ファイルフィルタを初期化する
'coXlonFltrIndx = 1～ : ファイルフィルタの初期選択
    If myXobjFileDialog Is Nothing Then Exit Sub
  Dim myXstrTmp As String, L As Long, U As Long
    On Error GoTo ExitPath
    L = LBound(myZstrAddFltr, 1): U = UBound(myZstrAddFltr, 1)
    myXstrTmp = myZstrAddFltr(L, L)
    On Error GoTo 0
    If UBound(myZstrAddFltr, 2) <> L + 1 Then Exit Sub
  Dim i As Long, myXstrFltrFile As String, myXstrFltrExtsn As String, n As Long
    If coXbisFltrClr = False Then
        n = myXobjFileDialog.Filters.Count
    Else
        myXobjFileDialog.Filters.Clear: n = 0
    End If
    On Error Resume Next
    For i = L To U
        myXstrFltrFile = Empty: myXstrFltrExtsn = Empty
        myXstrFltrFile = myZstrAddFltr(i, L + 0)
        myXstrFltrExtsn = myZstrAddFltr(i, L + 1)
        If myXstrFltrExtsn <> "" Then
            n = n + 1
            Call myXobjFileDialog.Filters.Add(myXstrFltrFile, myXstrFltrExtsn, n)
        End If
    Next i
    On Error GoTo 0
    If coXlonFltrIndx < 1 Or coXlonFltrIndx > n Then coXlonFltrIndx = 1
    myXobjFileDialog.FilterIndex = coXlonFltrIndx
ExitPath:
End Sub

 '定型Ｐ_ファイル選択用のFileDialogオブジェクトの初期状態を設定する
Private Sub PfixSetFileDialogForFileInitial( _
            myXobjFileDialog As Object, _
            Optional ByVal coXlonIniView As Long = msoFileDialogViewList, _
            Optional ByVal coXstrDfltPath As String = "", _
            Optional ByVal coXbisMultSlct As Boolean = False)
'coXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
'coXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
'coXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
'coXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
'coXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
'coXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
'coXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
'coXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
'coXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
'coXstrDfltPath : ファイルパス指定→ダイアログボックスにファイル名を表示
'coXstrDfltPath : フォルダパス指定→デフォルトフォルダを表示
'coXbisMultSlct = False : 複数のファイルを選択不可能
'coXbisMultSlct = True  : 複数のファイルを選択可能
    If myXobjFileDialog Is Nothing Then Exit Sub
    On Error Resume Next
    With myXobjFileDialog
        .InitialView = coXlonIniView
        .InitialFileName = coXstrDfltPath
        .AllowMultiSelect = coXbisMultSlct
    End With
    On Error GoTo 0
End Sub

 '定型Ｐ_ファイル選択用のFileDialogオブジェクトを使用してファイルを選択する
Private Sub PfixSelectFileByFileDialog( _
            myXlonFileCnt As Long, myZstrFilePath() As String, _
            ByVal myXobjFileDialog As Object)
'myZstrFilePath(i) : ファイルパス
    myXlonFileCnt = Empty: Erase myZstrFilePath
    If myXobjFileDialog Is Nothing Then Exit Sub
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXvarTmpPath As Variant, n As Long: n = L - 1
  Dim myXlonFDType As Long
    With myXobjFileDialog
        myXlonFDType = .DialogType
        If .Show = False Then Exit Sub
        myXlonFileCnt = .SelectedItems.Count
        ReDim myZstrFilePath(myXlonFileCnt) As String
        For Each myXvarTmpPath In .SelectedItems
            n = n + 1: myZstrFilePath(n) = CStr(myXvarTmpPath)
            If myXlonFDType = msoFileDialogOpen Or _
                myXlonFDType = msoFileDialogSaveAs Then .Execute
        Next
    End With
End Sub

 '定型Ｐ_ファイル一覧からベース名が複数の指定条件と一致するファイルパスを指定条件リストごとに抽出する
Private Sub PincPickUpBaseNameMatchFilePathMatrixSomeConditions( _
            myXlonOrdrCnt As Long, myXlonPckUpMax As Long, _
            myZlonBsMtchFileCnt() As Long, myZstrBsMtchFilePath() As String, _
            ByRef myZstrOrgFilePath() As String, ByRef myZvarOdrTrgtWrdPos() As Variant)
'Includes PfncbisCheckFileBaseNameSomeConditions
'myXlonOrdrCnt : ベース名条件数
'myXlonPckUpMax : 抽出した最大ファイル数
'myZlonBsMtchFileCnt(i, 1)  : 各ベース名条件での抽出ファイルパス数
'myZstrBsMtchFilePath(i, j) : 抽出ファイルパス
'myZstrOrgFilePath(k) : 元ファイルパス
'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
    myXlonOrdrCnt = Empty: myXlonPckUpMax = Empty
    Erase myZlonBsMtchFileCnt: Erase myZstrBsMtchFilePath
  Dim La As Long, Ua As Long, Lb As Long, Ub1 As Long, Ub2 As Long, Ub3 As Long
    On Error GoTo ExitPath
    La = LBound(myZstrOrgFilePath): Ua = UBound(myZstrOrgFilePath)
    Lb = LBound(myZvarOdrTrgtWrdPos, 1)
    Ub1 = UBound(myZvarOdrTrgtWrdPos, 1): Ub2 = UBound(myZvarOdrTrgtWrdPos, 2)
    Ub3 = UBound(myZvarOdrTrgtWrdPos, 3)
    On Error GoTo 0
    If Ub1 < Lb Or Ub3 <> Lb + 1 Then Exit Sub
'//抽出したい順番にベース名条件の検索を実行
  Dim myZlonTmp(1) As Long, Lc As Long: Lc = LBound(myZlonTmp)
  Dim Uc As Long: Uc = Ub1 + (Lc - Lb)
    ReDim myZlonBsMtchFileCnt(Uc, Lc) As Long
    ReDim myZstrBsMtchFilePath(Uc, Lc) As String
  Dim myZvarTrgtWrdPos() As Variant
  Dim myXstrFilePath As String, myXbisBsChck As Boolean, myZstrTmpPath() As String
  Dim m As Long, n As Long, Un As Long, myXlonNMax As Long, myXlonUnMax As Long
  Dim i As Long, j As Long, k As Long, p As Long
    m = 0: myXlonNMax = 0: myXlonUnMax = Lc
    For i = Lb To Ub1
   '//ファイル一覧からベース名が指定ファイル名と一致するファイルパスを抽出
        ReDim myZvarTrgtWrdPos(Ub2, Ub3) As Variant
        n = 0: Un = Lc - 1: Erase myZstrTmpPath
        For p = Lb To Ub2
            myZvarTrgtWrdPos(p, Lb + 0) = myZvarOdrTrgtWrdPos(i, p, Lb + 0)
            myZvarTrgtWrdPos(p, Lb + 1) = myZvarOdrTrgtWrdPos(i, p, Lb + 1)
        Next p
        For k = La To Ua
            myXstrFilePath = CStr(myZstrOrgFilePath(k))
            myXbisBsChck = PfncbisCheckFileBaseNameSomeConditions(myXstrFilePath, myZvarTrgtWrdPos)
            If myXbisBsChck = True Then
                n = n + 1: Un = Un + 1: ReDim Preserve myZstrTmpPath(Un) As String
                myZstrTmpPath(Un) = myXstrFilePath
            End If
        Next k
        If n > myXlonNMax Then myXlonNMax = n
        If Un > myXlonUnMax Then myXlonUnMax = Un
    '//抽出したファイルパスを出力配列変数に格納
        myZlonBsMtchFileCnt(i, Lc) = n
        If n <= 0 Then GoTo NextPath
        m = m + 1
        ReDim Preserve myZstrBsMtchFilePath(Uc, myXlonUnMax) As String
        For j = LBound(myZstrTmpPath) To UBound(myZstrTmpPath)
            myZstrBsMtchFilePath(i, j) = myZstrTmpPath(j)
        Next j
NextPath:
    Next i
    myXlonOrdrCnt = m: myXlonPckUpMax = myXlonNMax
ExitPath:
End Sub

 '定型Ｆ_指定ファイルのベース名が複数の指定条件と一致することを確認する
Private Function PfncbisCheckFileBaseNameSomeConditions( _
            ByVal myXstrFilePath As String, _
            ByRef myZvarTrgtWrdPos() As Variant) As Boolean
'myZvarTrgtWrdPos(i, 1) = i : 指定文字列
'myZvarTrgtWrdPos(i, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'myZvarTrgtWrdPos(i, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'myZvarTrgtWrdPos(i, 2) = 3 : 指定文字列をベースファイル名内に含む
    PfncbisCheckFileBaseNameSomeConditions = False
    If myXstrFilePath = "" Then Exit Function
  Dim L As Long, U1 As Long, U2 As Long
    On Error GoTo ExitPath
    L = LBound(myZvarTrgtWrdPos, 1)
    U1 = UBound(myZvarTrgtWrdPos, 1): U2 = UBound(myZvarTrgtWrdPos, 2)
    On Error GoTo 0
    If U1 < L Or U2 <> L + 1 Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExist As Boolean, myXstrOrgBaseName As String
    With myXobjFSO
        myXbisFileExist = .FileExists(myXstrFilePath)
        If myXbisFileExist = False Then Exit Function
        myXstrOrgBaseName = .getBaseName(myXstrFilePath)
    End With
    Set myXobjFSO = Nothing
  Dim i As Long, myXstrTrgtWrd As String, myXlonPosOptn As Long
    For i = LBound(myZvarTrgtWrdPos, 1) To UBound(myZvarTrgtWrdPos, 1)
        Call PforExtractArrayDataOfTargetWordPosition( _
                myXstrTrgtWrd, myXlonPosOptn, myZvarTrgtWrdPos, i, L)
        If myXstrTrgtWrd = "" Then GoTo NextPath
        If myXlonPosOptn = 1 Then
            If Left(myXstrOrgBaseName, Len(myXstrTrgtWrd)) <> myXstrTrgtWrd Then Exit Function
        ElseIf myXlonPosOptn = 2 Then
            If Right(myXstrOrgBaseName, Len(myXstrTrgtWrd)) <> myXstrTrgtWrd Then Exit Function
        ElseIf myXlonPosOptn = 3 Then
            If InStr(myXstrOrgBaseName, myXstrTrgtWrd) = 0 Then Exit Function
        Else
            Exit Function
        End If
NextPath:
    Next i
    PfncbisCheckFileBaseNameSomeConditions = True
ExitPath:
End Function
Private Sub PforExtractArrayDataOfTargetWordPosition( _
            myXstrTrgtWrd As String, myXlonPosOptn As Long, _
            ByRef myZvarTrgtWrdPos() As Variant, ByVal i As Long, L As Long)
'//配列変数からデータを抽出
    myXstrTrgtWrd = Empty: myXlonPosOptn = Empty
    On Error GoTo ExitPath
    myXstrTrgtWrd = CStr(myZvarTrgtWrdPos(i, L + 0))
    myXlonPosOptn = CLng(myZvarTrgtWrdPos(i, L + 1))
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXstrTrgtWrd = Empty: myXlonPosOptn = Empty
End Sub

 '定型Ｐ_指定2次元ファイルパス一覧のファイル名一覧を取得する(FileSystemObject使用)
Private Sub PfixGetFileName2DArrayByFSO( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZstrFileName() As String, _
            ByRef myZstrFilePath() As String)
'myZstrFileName(i, j) : ファイル名
'myZstrFilePath(i, j) : ファイルパス
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZstrFileName
  Dim myXstrTmp As String, L As Long, U1 As Long, U2 As Long
    On Error GoTo ExitPath
    L = LBound(myZstrFilePath, 1)
    U1 = UBound(myZstrFilePath, 1): U2 = UBound(myZstrFilePath, 2)
    myXstrTmp = myZstrFilePath(L, L)
    On Error GoTo 0
    myXlonRowCnt = U1 - L + 1: myXlonColCnt = U2 - L + 1
    ReDim myZstrFileName(U1, U2) As String
  Dim i As Long, j As Long, myXstrTmpPath As String, myXstrTmpName As String
  Dim myXbisFileExist As Boolean
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath, 1) To UBound(myZstrFilePath, 1)
        For j = LBound(myZstrFilePath, 2) To UBound(myZstrFilePath, 2)
            myXstrTmpPath = Empty: myXstrTmpName = Empty
            myXstrTmpPath = myZstrFilePath(i, j)
            myXbisFileExist = myXobjFSO.FileExists(myXstrTmpPath)
            If myXbisFileExist = True Then myXstrTmpName = myXobjFSO.getFileName(myXstrTmpPath)
            myZstrFileName(i, j) = myXstrTmpName
        Next j
    Next i
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '定型Ｐ_クリップボードにデータを格納する
Private Sub PfixDataPutInClipboard(ByVal myXvarInptData As Variant)
    If myXvarInptData = "" Then Exit Sub
  Dim myXdobClpBrd As New DataObject, myXvarClpBrdData As Variant
    With myXdobClpBrd
        .SetText myXvarInptData         '変数をDataObjectに格納
        .PutInClipboard                 'DataObjectのデータをクリップボードに格納
        .GetFromClipboard               'クリップボードからDataObjectにデータを取得
        myXvarClpBrdData = .GetText     'DataObjectのデータを変数に格納
    End With
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_指定文字列を含むファイル名のファイルを繰返し選択してそのパスを取得する
'Private Sub instCSlctFilePathRpt()
'  Dim Lc As Long
'  Dim myXlonFileCnt As Long, _
'        myZstrFileName() As String, myZstrFilePath() As String
'    'myZstrFileName(i, 1) : ファイル名
'    'myZstrFilePath(i, 1) : ファイルパス
''//クラス内変数への入力
'  Dim myXstrDfltFldrPath As String
'    myXstrDfltFldrPath = "1"
'    'myXstrDfltFldrPath = ""  : デフォルトパス指定無し
'    'myXstrDfltFldrPath = "C" : Cドライブをデフォルトパスに指定
'    'myXstrDfltFldrPath = "1" : このブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "2" : アクティブブックの親フォルダをデフォルトパスに指定
'    'myXstrDfltFldrPath = "*" : デフォルトパスを指定
'  Dim myXstrDfltFilePath As String
'    myXstrDfltFilePath = "1"
'    'myXstrDfltFilePath = ""  : デフォルトパス指定無し
'    'myXstrDfltFilePath = "1" : このブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "2" : アクティブブックをデフォルトパスに指定
'    'myXstrDfltFilePath = "*" : デフォルトパスを指定
'  Dim myXstrExtsn As String
'    myXstrExtsn = "PDF"
'  Dim myZstrAddFltr() As String: ReDim myZstrAddFltr(1, 2) As String
'    myZstrAddFltr(1, 1) = "PDFファイル"
'    myZstrAddFltr(1, 2) = "*.pdf"
'    'myZstrAddFltr(i, 1) : ファイルの候補を指定する文字列(ファイル)
'    'myZstrAddFltr(i, 2) : ファイルの候補を指定する文字列(フィルタ文字列)
'  Dim myXbisFltrClr As Boolean
'    myXbisFltrClr = False
'    'myXbisFltrClr = False : ファイルフィルタを初期化せずに追加する
'    'myXbisFltrClr = True  : ファイルフィルタを初期化する
'  Dim myXlonFltrIndx As Long
'    myXlonFltrIndx = 1
'    'myXlonFltrIndx = 1～ : ファイルフィルタの初期選択
'  Dim myXlonIniView As Long
'    myXlonIniView = msoFileDialogViewList
'    'myXlonIniView = msoFileDialogViewDetails    : ファイルを詳細情報と共に一覧表示
'    'myXlonIniView = msoFileDialogViewLargeIcons : ファイルを大きいアイコンで表示
'    'myXlonIniView = msoFileDialogViewList       : ファイルを詳細情報なしで一覧表示
'    'myXlonIniView = msoFileDialogViewPreview    : ファイルの一覧を表示し、選択したファイルをプレビュー ウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewProperties : ファイルの一覧を表示し、選択したファイルのプロパティをウィンドウ枠に表示
'    'myXlonIniView = msoFileDialogViewSmallIcons : ファイルを小さいアイコンで表示
'    'myXlonIniView = msoFileDialogViewThumbnail  : ファイルを縮小表示
'    'myXlonIniView = msoFileDialogViewTiles      : ファイルをアイコンで並べて表示
'    'myXlonIniView = msoFileDialogViewWebView    : ファイルを Web 表示
'  Dim myXbisMultSlct As Boolean
'    myXbisMultSlct = False
'    'myXbisMultSlct = False : 複数のファイルを選択不可能
'    'myXbisMultSlct = True  : 複数のファイルを選択可能
'  Dim myXlonOrdrCnt As Long, myXlonTrgtWrdCnt As Long, myZvarOdrTrgtWrdPos() As Variant
'    myXlonOrdrCnt = 2
'    myXlonTrgtWrdCnt = 2
'    ReDim myZvarOdrTrgtWrdPos(myXlonOrdrCnt, myXlonTrgtWrdCnt, 2) As Variant
'    myZvarOdrTrgtWrdPos(1, 1, 1) = "f"
'    myZvarOdrTrgtWrdPos(1, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(1, 2, 1) = "d"
'    myZvarOdrTrgtWrdPos(1, 2, 2) = 3
'    myZvarOdrTrgtWrdPos(2, 1, 1) = ""
'    myZvarOdrTrgtWrdPos(2, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(2, 2, 1) = "操作"
'    myZvarOdrTrgtWrdPos(2, 2, 2) = 3
'    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i番目の抽出ファイルの指定文字列:条件p
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : 指定文字列をベースファイル名の先頭に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : 指定文字列をベースファイル名の接尾に含む
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : 指定文字列をベースファイル名内に含む
''//CSlctFilePathRptクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim k As Long
'  Dim myXinsFilePathRpt As CSlctFilePathRpt: Set myXinsFilePathRpt = New CSlctFilePathRpt
'    With myXinsFilePathRpt
'    '//クラス内変数への入力
'        .letDfltFldrPath = myXstrDfltFldrPath
'        .letDfltFilePath = myXstrDfltFilePath
'        .letExtsn = myXstrExtsn
'        .letAddFltr = myZstrAddFltr
'        .letFltrClr = myXbisFltrClr
'        .letFltrIndx = myXlonFltrIndx
'        .letIniView = myXlonIniView
'        .letMultSlct = myXbisMultSlct
'        .letOdrTrgtWrdPosAry = myZvarOdrTrgtWrdPos
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXlonFileCnt = .getFileCnt
'        If myXlonFileCnt <= 0 Then GoTo JumpPath
'        k = myXlonFileCnt + Lo - 1
'        ReDim myZstrFileName(k, Lo) As String
'        ReDim myZstrFilePath(k, Lo) As String
'        Lc = .getOptnBase
'        For k = 1 To myXlonFileCnt
'            myZstrFileName(k + Lo - 1, Lo) = .getFileNameAry(k + Lc - 1)
'            myZstrFilePath(k + Lo - 1, Lo) = .getFilePathAry(k + Lc - 1)
'        Next k
'    End With
'JumpPath:
'    Set myXinsFilePathRpt = Nothing
'    Call variablesOfCSlctFilePathRpt(myXlonFileCnt, myZstrFileName) 'Debug.Print
'End Sub
'Private Sub variablesOfCSlctFilePathRpt( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CSlctFilePathRptクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
