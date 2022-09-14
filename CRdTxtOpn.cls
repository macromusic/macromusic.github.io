VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRdTxtOpn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfnclonCheckExcelBookOpening
'Includes PfixGetFileNameInformationByFSO
'Includes PfncobjOpenTextOnExcel
'Includes PfncobjGetExcelBookIfOpened
'Includes PfncobjLastCell
'Includes PfixGetSheetRangeData
'Includes PfixCloseFile

Option Explicit
Option Base 1

'◆ClassProc名_ファイルを開いてテキストファイルの内容を取得する
'Rev.002
  
'//モジュールメモリ
  Private Const MstrMdlName As String = "CRdTxtOpn"
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXstrDirPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String
  Private myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
            myZstrTxtData() As String
  
'//PropertySetLet_入力データ
  Private myXstrFilePath As String
  Private myXlonBgn As Long, myXlonEnd As Long
  Private myXbisSpltOptn As Boolean, myXstrSpltChr As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  Private myXbisTab As Boolean, myXbisSemicolon As Boolean, _
            myXbisComma As Boolean, myXbisSpace As Boolean, _
            myXbisOther As Boolean
  
'//モジュール内変数_データ
  Private myXlonFileOpndType As Long
  Private myXobjOpndFile As Object
  Private myXobjSheet As Object, _
            myXobjFrstCell As Object, myXobjLastCell As Object

'SubP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisTab = False: myXbisSemicolon = False
    myXbisComma = False: myXbisSpace = False
    myXbisOther = False
    
    myXlonFileOpndType = Empty
    Set myXobjOpndFile = Nothing
    Set myXobjSheet = Nothing
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getDirPath() As String
    getDirPath = myXstrDirPath
End Property

Public Property Get getFileName() As String
    getFileName = myXstrFileName
End Property

Public Property Get getBaseName() As String
    getBaseName = myXstrBaseName
End Property

Public Property Get getExtsn() As String
    getExtsn = myXstrExtsn
End Property

Public Property Get getTxtRowCnt() As Long
    getTxtRowCnt = myXlonTxtRowCnt
End Property

Public Property Get getTxtColCnt() As Long
    getTxtColCnt = myXlonTxtColCnt
End Property

Public Property Get getTxtDataAry(ByVal i As Long, ByVal j As Long) As String
    getTxtDataAry = myZstrTxtData(i, j)
'myZstrTxtData(i, j) : テキストファイル内容
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXstrDirPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    myXlonTxtRowCnt = Empty: myXlonTxtColCnt = Empty: Erase myZstrTxtData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Let letFilePath(ByVal RHS As String)
    myXstrFilePath = RHS
End Property

Public Property Let letRdBgnEnd(ByVal i As Long, ByVal RHS As Long)
    Select Case i
        Case 1: myXlonBgn = RHS
        Case 2: myXlonEnd = RHS
    End Select
'letRdBgnEnd(1) : 最初のデータ位置
'letRdBgnEnd(2) : 最後のデータ位置
End Property

Public Property Let letSpltOptn(ByVal RHS As Boolean)
    myXbisSpltOptn = RHS
'myXbisSpltOptn = True  : 文字列を分割処理する
'myXbisSpltOptn = False : 文字列を分割処理しない
End Property

Public Property Let letSpltChr(ByVal RHS As String)
    myXstrSpltChr = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'//指定ファイルが既に開いているか確認
    myXlonFileOpndType = PfnclonCheckExcelBookOpening(myXstrFilePath)
    If myXlonFileOpndType <= 0 And myXlonFileOpndType >= 4 Then GoTo ExitPath
    
'//ファイル名情報を取得
    Call PfixGetFileNameInformationByFSO( _
            myXstrDirPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrFilePath)
    
'//分割処理する文字列を設定
    If myXbisSpltOptn = False Then Exit Sub
    If myXstrSpltChr <> "" Then
        myXbisOther = True
    Else
        Select Case LCase(myXstrExtsn)
            Case "txt": myXbisTab = True
            Case "csv": myXbisComma = True
        End Select
    End If
    
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

'//指定ファイルを開く
    If myXlonFileOpndType = 1 Then
    '//指定ファイルが開いていない場合の処理
        Set myXobjOpndFile = PfncobjOpenTextOnExcel( _
                                myXstrFilePath, 1, xlDelimited, False, _
                                myXbisTab, myXbisSemicolon, myXbisComma, _
                                myXbisSpace, myXbisOther, myXstrSpltChr, False)
        If myXobjOpndFile Is Nothing Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
        
    ElseIf myXlonFileOpndType = 2 Then
    '//指定ファイルと同一名の別ファイルが開いている場合の処理
        Call PfixCloseFile(myXbisExitFlag, myXstrFileName)
        If myXbisExitFlag = True Then GoTo ExitPath
        Set myXobjOpndFile = PfncobjOpenTextOnExcel( _
                                myXstrFilePath, 1, xlDelimited, False, _
                                myXbisTab, myXbisSemicolon, myXbisComma, _
                                myXbisSpace, myXbisOther, myXstrSpltChr, False)
        If myXobjOpndFile Is Nothing Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
        
    ElseIf myXlonFileOpndType = 3 Then
    '//指定ファイルが開いている場合の処理
        Set myXobjOpndFile = PfncobjGetExcelBookIfOpened(myXstrFileName)
        If myXobjOpndFile Is Nothing Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
        
    End If

'//シート上の指定範囲またはデータ全範囲のデータを配列に取込む
    Set myXobjSheet = myXobjOpndFile.Worksheets(1)
    If myXlonEnd >= myXlonBgn Then
      Dim myXlonLastCol As Long
        myXlonLastCol = PfncobjLastCell(myXobjSheet).Column
        Set myXobjFrstCell = myXobjSheet.Cells(myXlonBgn, 1)
        Set myXobjLastCell = myXobjSheet.Cells(myXlonEnd, myXlonLastCol)
    End If
    Call PfixGetSheetRangeData( _
            myXlonTxtRowCnt, myXlonTxtColCnt, myZstrTxtData, _
            myXobjSheet, myXobjFrstCell, myXobjLastCell)
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag

'//指定ファイルを閉じる
    myXstrFileName = myXobjOpndFile.Name
    Call PfixCloseFile(myXbisExitFlag, myXstrFileName)
    If myXbisExitFlag = True Then Exit Sub
'    Debug.Print "PassFlag: " & meMstrMdlName & "6"     'PassFlag

ExitPath:
    Call initializeModuleVariables
End Sub


'===============================================================================================

 '定型Ｆ_指定エクセルブックが既に開いているか確認する
Private Function PfnclonCheckExcelBookOpening( _
            ByVal myXstrFullName As String) As Long
'PfnclonCheckExcelBookOpening = 0 : 指定ブックが存在しない
'PfnclonCheckExcelBookOpening = 1 : 開いていない
'PfnclonCheckExcelBookOpening = 2 : 指定ブックと同一名の別ブックが開いている
'PfnclonCheckExcelBookOpening = 3 : 指定ブックが開いている
    PfnclonCheckExcelBookOpening = Empty
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If myXobjFSO.FileExists(myXstrFullName) = False Then Exit Function
  Dim myXstrBookName As String
    myXstrBookName = myXobjFSO.getFileName(myXstrFullName)
    On Error GoTo ExitPath
  Dim myXstrTmp As String: myXstrTmp = Workbooks(myXstrBookName).FullName
    On Error GoTo 0
    If myXstrTmp = myXstrFullName Then
        PfnclonCheckExcelBookOpening = 3
    Else
        PfnclonCheckExcelBookOpening = 2
    End If
    Set myXobjFSO = Nothing
    Exit Function
ExitPath:
    PfnclonCheckExcelBookOpening = 1
    Set myXobjFSO = Nothing
End Function

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
        myXstrFileName = .getFileName(myXstrFilePath)           'ファイル名
        myXstrBaseName = .getBaseName(myXstrFilePath)           'ファイルベース名
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         'ファイル拡張子
    End With
    Set myXobjFSO = Nothing
End Sub

 '定型Ｆ_エクセルシート上にテキストファイルを開く
Private Function PfncobjOpenTextOnExcel( _
            ByVal myXstrFilePath As String, _
            Optional ByVal coXlonStartRow As Long = 1, _
            Optional ByVal coXstrDataType As String = xlDelimited, _
            Optional ByVal coXbisConsecutiveDelimiter As Boolean = False, _
            Optional ByVal coXbisTab As Boolean = False, _
            Optional ByVal coXbisSemicolon As Boolean = False, _
            Optional ByVal coXbisComma As Boolean = False, _
            Optional ByVal coXbisSpace As Boolean = False, _
            Optional ByVal coXbisOther As Boolean = False, _
            Optional ByVal coXstrOtherChar As String = "", _
            Optional ByVal coXlonFieldInfo As Boolean = False) As Object
'DataType:=xlDelimited  = 1 : 区切り文字で区切る
'DataType:=xlFixedWidth = 2 : ファイルのデータが固定幅の列に配置
'ConsecutiveDelimiter:=False : 連続した区切り文字を1文字として扱わない
'ConsecutiveDelimiter:=True  : 連続した区切り文字を1文字として扱う
'FieldInfo:=False : 各列のデータ形式を変更しない
'FieldInfo:=True  : 各列のデータ形式を文字列に指定
    Set PfncobjOpenTextOnExcel = Nothing
  Dim myXlonColCnt As Long: myXlonColCnt = ActiveSheet.Columns.Count
  Dim myXvarFldInfo As Variant: myXvarFldInfo = Array(1, 1)
'  Dim myXvarFldInfo() As Variant: ReDim myXvarFldInfo(myXlonColCnt) As Variant
'  Dim i As Long
'    If coXlonFieldInfo = False Then
'        For i = 1 To myXlonColCnt
'            myXvarFldInfo(i) = Array(i, xlGeneralFormat)    '全フィールドを一般形式に指定
'        Next i
'    Else
'        For i = 1 To myXlonColCnt
'            myXvarFldInfo(i) = Array(i, xlTextFormat)       '全フィールドをテキスト形式に指定
'        Next i
'    End If
    On Error Resume Next
    Workbooks.OpenText Filename:=myXstrFilePath, _
                        StartRow:=coXlonStartRow, _
                        DataType:=coXstrDataType, _
                        ConsecutiveDelimiter:=coXbisConsecutiveDelimiter, _
                        Tab:=coXbisTab, _
                        Semicolon:=coXbisSemicolon, _
                        Comma:=coXbisComma, _
                        Space:=coXbisSpace, _
                        Other:=coXbisOther, _
                        OtherChar:=coXstrOtherChar, _
                        FieldInfo:=myXvarFldInfo
    On Error GoTo 0
    Set PfncobjOpenTextOnExcel = ActiveWorkbook
    Erase myXvarFldInfo
End Function

 '定型Ｆ_指定名のエクセルブックが既に開いていればブックオブジェクトを取得する
Private Function PfncobjGetExcelBookIfOpened( _
            ByVal myXstrBookName As String) As Object
    Set PfncobjGetExcelBookIfOpened = Nothing
    On Error GoTo ExitPath
    Set PfncobjGetExcelBookIfOpened = Workbooks(myXstrBookName)
    On Error GoTo 0
ExitPath:
End Function

 '定型Ｆ_シート上のデータ記載範囲の最後のセルを取得する
Private Function PfncobjLastCell(ByVal myXobjSheet As Object) As Object
    Set PfncobjLastCell = Nothing
    If myXobjSheet Is Nothing Then Exit Function
    With myXobjSheet
        Set PfncobjLastCell = .Cells.SpecialCells(xlCellTypeLastCell)
    End With
End Function

 '定型Ｐ_シート上の指定範囲またはデータ全範囲のデータを配列に取込む
Private Sub PfixGetSheetRangeData( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZstrShtData() As String, _
            myXobjSheet As Object, myXobjFrstCell As Object, myXobjLastCell As Object)
'myZstrShtData(i, j) : データ
    myXlonRowCnt = Empty: myXlonColCnt = Empty: Erase myZstrShtData
    If myXobjSheet Is Nothing Then Exit Sub
'//シート上の指定範囲をオブジェクト配列に取込む
  Dim myXobjShtRng As Object
    If myXobjFrstCell Is Nothing Then Set myXobjFrstCell = myXobjSheet.Cells(1, 1)
    If myXobjLastCell Is Nothing Then _
        Set myXobjLastCell = myXobjSheet.Cells.SpecialCells(xlCellTypeLastCell)
    Set myXobjShtRng = myXobjSheet.Range(myXobjFrstCell, myXobjLastCell)
    myXlonRowCnt = myXobjShtRng.Rows.Count
    myXlonColCnt = myXobjShtRng.Columns.Count
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then Exit Sub
'//オブジェクト配列からデータを取得
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZvarRngData As Variant, i As Long, j As Long
    i = myXlonRowCnt - (1 - L): j = myXlonColCnt - (1 - L)
    ReDim myZstrShtData(i, j) As String
    myZvarRngData = myXobjShtRng.Value
    On Error Resume Next
    If myXlonRowCnt = 1 And myXlonColCnt = 1 Then
        myZstrShtData(L, L) = myZvarRngData
    Else
        For j = LBound(myZvarRngData, 2) To UBound(myZvarRngData, 2)
            For i = LBound(myZvarRngData, 1) To UBound(myZvarRngData, 1)
                myZstrShtData(i, j) = myZvarRngData(i, j)
            Next i
        Next j
    End If
    On Error GoTo 0
    Set myXobjShtRng = Nothing: myZvarRngData = Empty
End Sub

 '定型Ｐ_指定ファイルを閉じる
Private Sub PfixCloseFile(myXbisExitFlag As Boolean, _
            ByVal myXstrFileName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Windows(myXstrFileName).Close
    On Error GoTo 0
    GoTo ExitPath
ErrPath:
    myXbisExitFlag = True
ExitPath:
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_ファイルを開いてテキストファイルの内容を取得する
'Private Sub instCRdTxtOpn()
'  Dim Lc As Long
'  Dim myXstrDirPath As String, myXstrFileName As String, _
'        myXstrBaseName As String, myXstrExtsn As String
'  Dim myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
'        myZstrTxtData() As String
'    'myZstrTxtData(i, j) : テキストファイル内容
''//クラス内変数への入力
'  Dim myXstrFilePath As String
'    myXstrFilePath = ""
'  Dim myXlonBgn As Long, myXlonEnd As Long
'    myXlonBgn = 1
'    myXlonEnd = 0
'  Dim myXbisSpltOptn As Boolean, myXstrSpltChr As String
'    'myXbisSpltOptn = True  : 文字列を分割処理する
'    'myXbisSpltOptn = False : 文字列を分割処理しない
'    myXbisSpltOptn = True
'    myXstrSpltChr = ""
''//CRdTxtOpnクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsRdTxtOpn As CRdTxtOpn: Set myXinsRdTxtOpn = New CRdTxtOpn
'    With myXinsRdTxtOpn
'    '//クラス内変数への入力
'    '//テキストファイルパスを指定
'        .letFilePath = myXstrFilePath
'        .letRdBgnEnd(1) = myXlonBgn
'        .letRdBgnEnd(2) = myXlonEnd
'        .letSpltOptn = myXbisSpltOptn
'        .letSpltChr = myXstrSpltChr
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXstrDirPath = .getDirPath
'        myXstrFileName = .getFileName
'        myXstrBaseName = .getBaseName
'        myXstrExtsn = .getExtsn
'        myXlonTxtRowCnt = .getTxtRowCnt
'        myXlonTxtColCnt = .getTxtColCnt
'        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
'        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
'        ReDim myZstrTxtData(i, j) As String
'        Lc = .getOptnBase
'        For i = 1 To myXlonTxtRowCnt
'            For j = 1 To myXlonTxtColCnt
'                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
'            Next j
'        Next i
'    End With
'JumpPath:
'    Set myXinsRdTxtOpn = Nothing
'    Call variablesOfCRdTxtOpn(myXlonTxtRowCnt, myZstrTxtData) 'Debug.Print
'End Sub
'Private Sub variablesOfCRdTxtOpn( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CRdTxtOpnクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
