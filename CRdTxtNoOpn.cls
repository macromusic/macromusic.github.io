VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRdTxtNoOpn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckFileExist
'Includes PfnclonCountTextLineByLineInput
'Includes PfixGetFileNameInformationByFSO
'Includes PfixReadFullTextLineWithoutFileOpening
'Includes PfixReadRangeTextLineWithoutFileOpening
'Includes PincSplitString2DArray
'Includes PfixStringSplit

Option Explicit
Option Base 1

'◆ClassProc名_ファイルを開かずにテキストファイルの内容を取得する
'Rev.008

'//モジュールメモリ
  Private Const meMstrMdlName As String = "CRdTxtNoOpn"
  
'//モジュール内定数
  Private Const coXstrTXT As String = "txt"
  Private Const coXstrCSV As String = "csv"
  
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
  
'//モジュール内変数_データ
  Private myXlonRdDataCnt As Long, myZstrRdData() As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRdDataCnt = Empty: Erase myZstrRdData
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

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
    
'//指定ファイルの存在を確認
    If PfncbisCheckFileExist(myXstrFilePath) = False Then GoTo ExitPath
    
'//ファイル名情報を取得
    Call PfixGetFileNameInformationByFSO( _
            myXstrDirPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrFilePath)
    
'//分割処理する文字列を設定
'Chr(9)  : タブ文字(vbTab)
'Chr(44) : カンマ(,)
    If myXbisSpltOptn = False Then Exit Sub
    If myXstrSpltChr <> "" Then Exit Sub
    Select Case LCase(myXstrExtsn)
        Case LCase(coXstrTXT): myXstrSpltChr = vbTab
        Case LCase(coXstrCSV): myXstrSpltChr = Chr(44)
    End Select
    
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
    
'//定型Ｆ_ファイルを開かずにテキストファイルからデータ行数を読み込む
  Dim myXlonTxtLineCnt As Long
    myXlonTxtLineCnt = PfnclonCountTextLineByLineInput(myXstrFilePath)
    
'//テキストファイルを開かずにデータを読み込み
    If myXlonEnd < myXlonBgn Or _
            myXlonBgn > myXlonTxtLineCnt Or myXlonEnd > myXlonTxtLineCnt Then
    '//全行データを読み込む
        Call PfixReadFullTextLineWithoutFileOpening( _
                myXlonRdDataCnt, myZstrRdData, _
                myXstrFilePath)
        
    Else
    '//指定範囲内の行データを読み込む
        Call PfixReadRangeTextLineWithoutFileOpening( _
                myXlonRdDataCnt, myZstrRdData, _
                myXstrFilePath, myXlonBgn, myXlonEnd)
        
    End If
    If myXlonRdDataCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//文字列を分割
    If myXbisSpltOptn = False Then
    '//分割しない
        myXlonTxtRowCnt = myXlonRdDataCnt
        myXlonTxtColCnt = 1
        myZstrTxtData = myZstrRdData
    Else
    '//分割する
        Call PincSplitString2DArray( _
                myXlonTxtRowCnt, myXlonTxtColCnt, myZstrTxtData, _
                myZstrRdData, myXstrSpltChr)
        If myXlonTxtRowCnt <= 0 Then GoTo ExitPath
    End If
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag

ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================
 
 '定型Ｆ_指定ファイルの存在を確認する
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｆ_ファイルを開かずにテキストファイルからデータ行数を読み込む
Private Function PfnclonCountTextLineByLineInput( _
            ByVal myXstrFilePath As String) As Long
    PfnclonCountTextLineByLineInput = Empty
  Dim myXlonFNo As Long: myXlonFNo = FreeFile
    On Error GoTo ExitPath
    Open myXstrFilePath For Input As #myXlonFNo
    On Error GoTo 0
  Dim myXvarTmp As Variant, n As Long: n = n + 1
    Do Until EOF(myXlonFNo)
        Line Input #myXlonFNo, myXvarTmp
        n = n + 1
    Loop
    Close #myXlonFNo
    PfnclonCountTextLineByLineInput = n
ExitPath:
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
  Dim myXbisFileExist As Boolean
    With myXobjFSO
        myXbisFileExist = .FileExists(myXstrFilePath)
        If myXbisFileExist = False Then Exit Sub
        myXstrPrntPath = .GetParentFolderName(myXstrFilePath)   '親フォルダパス
        myXstrFileName = .getFileName(myXstrFilePath)           'ファイル名
        myXstrBaseName = .getBaseName(myXstrFilePath)           'ファイルベース名
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         'ファイル拡張子
    End With
    Set myXobjFSO = Nothing
End Sub
 
 '定型Ｐ_ファイルを開かずにテキストファイルから1行づつ全行データを読み込む
Private Sub PfixReadFullTextLineWithoutFileOpening( _
            myXlonTxtDataCnt As Long, myZstrTxtData() As String, _
            ByVal myXstrFilePath As String)
'myZstrTxtData(i, 1) : テキストファイル内容
    myXlonTxtDataCnt = Empty: Erase myZstrTxtData
  Dim myXlonFNo As Long: myXlonFNo = FreeFile
    On Error GoTo ExitPath
    Open myXstrFilePath For Input As #myXlonFNo
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZstrTmp() As String, myXvarTmp As Variant, n As Long: n = L - 1
    On Error Resume Next
    Do Until EOF(myXlonFNo)
        Line Input #myXlonFNo, myXvarTmp
        n = n + 1: ReDim Preserve myZstrTmp(n) As String
        myZstrTmp(n) = CStr(myXvarTmp)
    Loop
    On Error GoTo 0
    Close #myXlonFNo
    myXlonTxtDataCnt = n - L + 1
    If myXlonTxtDataCnt <= 0 Then Exit Sub
    ReDim myZstrTxtData(n, L) As String
  Dim i As Long
    For i = LBound(myZstrTmp) To UBound(myZstrTmp)
        myZstrTxtData(i, L) = myZstrTmp(i)
    Next i
    Erase myZstrTmp
ExitPath:
End Sub
 
 '定型Ｐ_ファイルを開かずにテキストファイルから1行づつ指定範囲内の行データを読み込む
Private Sub PfixReadRangeTextLineWithoutFileOpening( _
            myXlonTxtDataCnt As Long, myZstrTxtData() As String, _
            ByVal myXstrFilePath As String, _
            Optional ByVal coXlonBgn As Long = 1, _
            Optional ByVal coXlonEnd As Long = 0)
'myZstrTxtData(i, 1) : テキストファイル内容
    myXlonTxtDataCnt = Empty: Erase myZstrTxtData
    If coXlonBgn <= 0 Then Exit Sub
  Dim myXlonFNo As Long: myXlonFNo = FreeFile
    On Error GoTo ExitPath
    Open myXstrFilePath For Input As #myXlonFNo
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZstrTmp() As String, myXvarTmp As Variant
  Dim m As Long, n As Long: m = 0: n = L - 1
    On Error Resume Next
    Do Until EOF(myXlonFNo)
        Line Input #myXlonFNo, myXvarTmp
        m = m + 1
        If m < coXlonBgn Then GoTo NextPath
        If m > coXlonEnd Then Exit Do
        n = n + 1: ReDim Preserve myZstrTmp(n) As String
        myZstrTmp(n) = CStr(myXvarTmp)
NextPath:
    Loop
    On Error GoTo 0
    Close #myXlonFNo
    myXlonTxtDataCnt = n - L + 1
    If myXlonTxtDataCnt <= 0 Then Exit Sub
    ReDim myZstrTxtData(n, L) As String
  Dim i As Long
    For i = LBound(myZstrTmp) To UBound(myZstrTmp)
        myZstrTxtData(i, L) = myZstrTmp(i)
    Next i
    Erase myZstrTmp
ExitPath:
End Sub

 '定型Ｐ_2次元の文字列配列内のデータを指定文字で分割する
Private Sub PincSplitString2DArray( _
            myXlonSpltDataRowCnt As Long, myXlonSpltDataColCnt As Long, _
            myZstrSpltData() As String, _
            ByRef myZstrOrgData() As String, ByVal myXstrSpltChr As String)
'Includes PfixSplitString
'myZstrSpltData(i, j) : 分割後データ配列
'myZstrOrgData(i, 1) : 元データ配列
    myXlonSpltDataRowCnt = Empty: myXlonSpltDataColCnt = Empty
    Erase myZstrSpltData
  Dim myXstrTmp As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgData, 1): myXstrTmp = myZstrOrgData(L, L)
    On Error GoTo 0
    myXlonSpltDataRowCnt = UBound(myZstrOrgData, 1) - LBound(myZstrOrgData, 1) + 1
    Dim r As Long
    r = myXlonSpltDataRowCnt + L - 1: ReDim myZstrSpltData(r, L) As String
  Dim myXlonSpltCnt As Long, myZstrSpltTmp() As String
  Dim i As Long, j As Long, k As Long: k = 0
    For i = LBound(myZstrOrgData, 1) To UBound(myZstrOrgData, 1)
        myXstrTmp = myZstrOrgData(i, L)
        Call PfixSplitString(myXlonSpltCnt, myZstrSpltTmp, myXstrTmp, myXstrSpltChr)
        If myXlonSpltCnt <= 0 Then GoTo NextPath
        If k < myXlonSpltCnt Then
            k = myXlonSpltCnt
            ReDim Preserve myZstrSpltData(r, k + L - 1) As String
        End If
        For j = LBound(myZstrSpltTmp) To UBound(myZstrSpltTmp)
            myZstrSpltData(i, j) = myZstrSpltTmp(j)
        Next j
NextPath:
    Next i
    myXlonSpltDataColCnt = k
    Erase myZstrSpltTmp
ExitPath:
End Sub

 '定型Ｐ_指定文字列を指定文字で分割して情報を抽出する
Private Sub PfixSplitString( _
            myXlonSpltCnt As Long, myZstrSpltData() As String, _
            ByVal myXstrOrgData As String, ByVal myXstrSpltChr As String)
'myZstrSpltData(i) : 分割情報
    myXlonSpltCnt = Empty: Erase myZstrSpltData
    If myXstrOrgData = "" Or myXstrSpltChr = "" Then Exit Sub
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myZvarTmp As Variant
    myZvarTmp = Split(myXstrOrgData, myXstrSpltChr)     '文字列を指定文字で分割
    myXlonSpltCnt = UBound(myZvarTmp) - LBound(myZvarTmp) + 1
    If myXlonSpltCnt <= 0 Then Exit Sub
  Dim i As Long, n As Long: n = L - 1
    i = myXlonSpltCnt + L - 1: ReDim myZstrSpltData(i) As String
    For i = LBound(myZvarTmp) To UBound(myZvarTmp)
        n = n + 1: myZstrSpltData(n) = CStr(myZvarTmp(i))
    Next i
    myZvarTmp = Empty
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_ファイルを開かずにテキストファイルの内容を取得する
'Private Sub instCRdTxtNoOpn()
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
''//CRdTxtNoOpnクラスモジュールのインスタンス生成
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim i As Long, j As Long
'  Dim myXinsRdTxtNoOpn As CRdTxtNoOpn: Set myXinsRdTxtNoOpn = New CRdTxtNoOpn
'    With myXinsRdTxtNoOpn
'    '//クラス内変数への入力
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
'    Set myXinsRdTxtNoOpn = Nothing
'    Call variablesOfCRdTxtNoOpn(myXlonTxtRowCnt, myZstrTxtData) 'Debug.Print
'End Sub
'Private Sub variablesOfCRdTxtNoOpn( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CRdTxtNoOpnクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim i As Long, j As Long
'    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
'        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
'            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
'        Next j
'    Next i
'End Sub
