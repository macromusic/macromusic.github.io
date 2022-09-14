VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckFileExist
'Includes PfixDeleteFile
'Includes PfixGetFileNameInformation

Option Explicit
Option Base 1

'◆ClassProc名_ファイルを削除する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CFileDelete"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonExeFileCnt As Long, _
            myZstrExeFilePath() As String, myZstrExeFileName() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisCheckOptn As Boolean
  
'//PropertySetLet_入力データ
  Private myXlonDeleteFileCnt As Long, myZstrDeleteFilePath() As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXstrDelFilePath As String
  Private myXstrExeFilePath As String, myXstrExeFileName As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrDelFilePath = Empty
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

Public Property Let letDeleteFileCnt(ByVal RHS As Long)
    myXlonDeleteFileCnt = RHS
End Property

Public Property Let letDeleteFilePathAry(ByRef RHS() As String)
'myZstrDeleteFilePath(i) : 元データ配列
    myZstrDeleteFilePath() = RHS()
End Property

Public Property Let letCheckOptn(ByVal RHS As Boolean)
    myXbisCheckOptn = RHS
'myXbisCheckOptn = True   : 削除直前に再確認する
'myXbisCheckOptn = False  : 削除確認無し
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If PfncbisCheckArrayDimension(myZstrDeleteFilePath, 1) = False Then GoTo ExitPath
    
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
    For Each myXvarTmpPath In myZstrDeleteFilePath
        myXstrDelFilePath = CStr(myXvarTmpPath)
 
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
 
'//指定ファイルの存在を確認
    If PfncbisCheckFileExist(myXstrDelFilePath) = False Then GoTo ExitPath

'//ファイルを削除
    Call PfixDeleteFile(myXbisExitFlag, myXstrDelFilePath, myXbisCheckOptn)
    If myXbisExitFlag = True Then GoTo ExitPath
    myXstrExeFilePath = myXstrDelFilePath
    
'//指定ファイルのファイル名情報を取得
  Dim myXstrPrntPath As String, myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformation( _
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

 '定型Ｆ_指定ファイルの存在を確認する
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_ファイルを削除する
Private Sub PfixDeleteFile(myXbisExitFlag As Boolean, _
            ByVal myXstrFilePath As String, _
            Optional ByVal coXbisMsgBxOptn As Boolean = False)
'coXbisMsgBxOptn = False : 削除確認無し
'coXbisMsgBxOptn = True  : 削除直前に再確認する
    myXbisExitFlag = False
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If myXobjFSO.FileExists(myXstrFilePath) = False Then GoTo ExitPath
    If coXbisMsgBxOptn = True Then
      Dim myXstrTitle As String, myXstrPrompt As String
        myXstrTitle = "ファイル削除の確認"
        myXstrPrompt = "本当にファイルを削除しますか？" _
                & vbCrLf & myXstrFilePath _
                & vbCrLf _
                & vbCrLf & "「OK」：削除" _
                & vbCrLf & "「キャンセル」：中止"
      Dim myXlonMsgBox As Long
        myXlonMsgBox = MsgBox(myXstrPrompt, vbOKCancel, myXstrTitle)
        If myXlonMsgBox = vbCancel Then GoTo ExitPath
    End If
    Call myXobjFSO.DeleteFile(myXstrFilePath)
    Set myXobjFSO = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

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

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_ファイルを削除する
'Private Sub instCFileDelete()
'  Dim myXbisCmpltFlag As Boolean
'  Dim Lc As Long
'  Dim myXlonExeFileCnt As Long, _
'        myZstrExeFilePath() As String, myZstrExeFileName() As String
'    'myZstrExeFilePath(k) : 実行ファイルパス
'    'myZstrExeFileName(k) : 実行ファイル名
''//クラス内変数への入力
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'  Dim i As Long
'  Dim myXlonDeleteFileCnt As Long, myZstrDeleteFilePath() As String
'    'myZstrDeleteFilePath(k) : 削除ファイルパス
'    myXlonDeleteFileCnt = 2
'    i = myXlonDeleteFileCnt + L - 1: ReDim myZstrDeleteFilePath(i) As String
'    myZstrDeleteFilePath(1) = ThisWorkbook.Path & "\test1" & ".txt"
'    myZstrDeleteFilePath(2) = ThisWorkbook.Path & "\test2" & ".txt"
''  Dim myXlonDeleteFileCnt As Long, myZstrDeleteFilePath() As String
''    'myZstrDeleteFilePath(k) : 削除ファイルパス
''    myXlonDeleteFileCnt = XarbLong
''    i = myXlonDeleteFileCnt + L - 1: ReDim myZstrDeleteFilePath(i) As String
''    myZstrDeleteFilePath(1) = XarbString
''//CFileDeleteクラスモジュールのインスタンス生成
'  Dim myXinsFileDelete As CFileDelete: Set myXinsFileDelete = New CFileDelete
'    With myXinsFileDelete
'    '//クラス内変数への入力
'        .letCheckOptn = True
'        If myXlonDeleteFileCnt <= 0 Then GoTo JumpPath
'        .letDeleteFileCnt = myXlonDeleteFileCnt
'        .letDeleteFilePathAry() = myZstrDeleteFilePath()
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
'    Set myXinsFileDelete = Nothing
'    Call variablesOfCFileDelete(myXlonExeFileCnt, myZstrExeFilePath)     'Debug.Print
'JumpPath:
'End Sub
'Private Sub variablesOfCFileDelete( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CFileDeleteクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
