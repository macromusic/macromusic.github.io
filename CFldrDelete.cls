VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFldrDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckFolderExist
'Includes PfixDeleteFolder
'Includes PfixGetFolderNameInformation

Option Explicit
Option Base 1

'◆ClassProc名_フォルダを削除する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CFldrDelete"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertyGet_出力データ
  Private Lc As Long
  Private myXlonExeFldrCnt As Long, _
            myZstrExeFldrPath() As String, myZstrExeFldrName() As String
  
'//PropertySetLet_入力制御信号
  Private myXbisCheckOptn As Boolean
  
'//PropertySetLet_入力データ
  Private myXlonDeleteFldrCnt As Long, myZstrDeleteFldrPath() As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXstrDelFldrPath As String
  Private myXstrExeFldrPath As String, myXstrExeFldrName As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrDelFldrPath = Empty
    myXstrExeFldrPath = Empty: myXstrExeFldrName = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getCmpltFlag() As Boolean
    getCmpltFlag = myXbisCmpltFlag
End Property

Public Property Get getOptnBase() As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    getOptnBase = L
End Property

Public Property Get getExeFldrCnt() As Long
    getExeFldrCnt = myXlonExeFldrCnt
End Property

Public Property Get getExeFldrPathAry(ByVal i As Long) As String
    getExeFldrPathAry = myZstrExeFldrPath(i)
'myZstrExeFldrPath(i) : 実行したファイルパス
End Property

Public Property Get getExeFldrNameAry(ByVal i As Long) As String
    getExeFldrNameAry = myZstrExeFldrName(i)
'myZstrExeFldrName(i) : 実行したファイル名
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonExeFldrCnt = Empty: Erase myZstrExeFldrPath: Erase myZstrExeFldrName
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letDeleteFldrCnt(ByVal RHS As Long)
    myXlonDeleteFldrCnt = RHS
End Property

Public Property Let letDeleteFldrPathAry(ByRef RHS() As String)
'myZstrDeleteFldrPath(i) : 元データ配列
    myZstrDeleteFldrPath() = RHS()
End Property

Public Property Let letCheckOptn(ByVal RHS As Boolean)
    myXbisCheckOptn = RHS
'myXbisCheckOptn = True   : 削除直前に再確認する
'myXbisCheckOptn = False  : 削除確認無し
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    If PfncbisCheckArrayDimension(myZstrDeleteFldrPath, 1) = False Then GoTo ExitPath
    
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
    
'//C:フォルダリストを順次実行
  Dim myXvarTmpPath As Variant, n As Long: n = 0
    For Each myXvarTmpPath In myZstrDeleteFldrPath
        myXstrDelFldrPath = CStr(myXvarTmpPath)
 
    '//P:各フォルダのデータ加工処理
        Call prsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "2-" & k  'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeFldrPath(n) As String
        ReDim Preserve myZstrExeFldrName(n) As String
        myZstrExeFldrPath(n) = myXstrExeFldrPath
        myZstrExeFldrName(n) = myXstrExeFldrName
NextPath:
    Next
    myXlonExeFldrCnt = n
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

'PrcsP_各フォルダのデータ加工処理
Private Sub prsProcForLoop()
    myXbisExitFlag = False

'//指定フォルダの存在を確認
  Dim myXstrDirPath As String
    If PfncbisCheckFolderExist(myXstrDelFldrPath) = False Then GoTo ExitPath
    
'//フォルダを削除
    Call PfixDeleteFolder(myXbisExitFlag, myXstrDelFldrPath, myXbisCheckOptn)
    If myXbisExitFlag = True Then GoTo ExitPath
    myXstrExeFldrPath = myXstrDelFldrPath
    
'//指定フォルダのフォルダ名情報を取得
  Dim myXstrPrntPath As String
    Call PfixGetFolderNameInformation( _
            myXstrPrntPath, myXstrExeFldrName, _
            myXstrExeFldrPath)
    If myXstrExeFldrName = "" Then GoTo ExitPath
    
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

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '定型Ｐ_フォルダを削除する
Private Sub PfixDeleteFolder(myXbisExitFlag As Boolean, _
            ByVal myXstrDirPath As String, _
            Optional ByVal coXbisCheckOptn As Boolean = True)
'coXbisCheckOptn = True   : 削除直前に再確認する
'coXbisCheckOptn = False  : 削除確認無し
    myXbisExitFlag = False
    If myXstrDirPath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If myXobjFSO.FolderExists(myXstrDirPath) = False Then GoTo ExitPath
    If coXbisCheckOptn = True Then
      Dim myXstrTitle As String, myXstrPrompt As String
        myXstrTitle = "フォルダ削除の確認"
        myXstrPrompt = "本当にフォルダを削除しますか？" _
                & vbCrLf & myXstrDirPath _
                & vbCrLf _
                & vbCrLf & "「OK」：削除" _
                & vbCrLf & "「キャンセル」：中止"
      Dim myXlonMsgBox As Long
        myXlonMsgBox = MsgBox(myXstrPrompt, vbOKCancel, myXstrTitle)
        If myXlonMsgBox = vbCancel Then GoTo ExitPath
    End If
    Call myXobjFSO.DeleteFolder(myXstrDirPath)
    Set myXobjFSO = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

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

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_フォルダを削除する
'Private Sub instCFldrDelete()
'  Dim myXbisCmpltFlag As Boolean
'  Dim Lc As Long
'  Dim myXlonExeFldrCnt As Long, _
'        myZstrExeFldrPath() As String, myZstrExeFldrName() As String
'    'myZstrExeFldrPath(k) : 実行フォルダパス
'    'myZstrExeFldrName(k) : 実行フォルダ名
''//クラス内変数への入力
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'  Dim i As Long
'  Dim myXlonDeleteFldrCnt As Long, myZstrDeleteFldrPath() As String
'    'myZstrDeleteFldrPath(k) : 削除フォルダパス
'    myXlonDeleteFldrCnt = 2
'    i = myXlonDeleteFldrCnt + L - 1: ReDim myZstrDeleteFldrPath(i) As String
'    myZstrDeleteFldrPath(1) = ThisWorkbook.Path & "\test1"
'    myZstrDeleteFldrPath(2) = ThisWorkbook.Path & "\test2"
''  Dim myXlonDeleteFldrCnt As Long, myZstrDeleteFldrPath() As String
''    'myZstrDeleteFldrPath(k) : 削除フォルダパス
''    myXlonDeleteFldrCnt = XarbLong
''    i = myXlonDeleteFldrCnt + L - 1: ReDim myZstrDeleteFldrPath(i) As String
''    myZstrDeleteFldrPath(1) = XarbString
''//CFldrDeleteクラスモジュールのインスタンス生成
'  Dim myXinsFldrDelete As CFldrDelete: Set myXinsFldrDelete = New CFldrDelete
'    With myXinsFldrDelete
'    '//クラス内変数への入力
'        .letCheckOptn = True
'        If myXlonDeleteFldrCnt <= 0 Then GoTo JumpPath
'        .letDeleteFldrCnt = myXlonDeleteFldrCnt
'        .letDeleteFldrPathAry() = myZstrDeleteFldrPath()
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXbisCmpltFlag = .getCmpltFlag
'        myXlonExeFldrCnt = .getExeFldrCnt
'        If myXlonExeFldrCnt <= 0 Then GoTo JumpPath
'        i = myXlonExeFldrCnt + L - 1
'        ReDim myZstrExeFldrPath(i) As String
'        ReDim myZstrExeFldrName(i) As String
'        Lc = .getOptnBase
'        For i = 1 To myXlonExeFldrCnt
'            myZstrExeFldrPath(i + L - 1) = .getExeFldrPathAry(i + Lc - 1)
'            myZstrExeFldrName(i + L - 1) = .getExeFldrNameAry(i + Lc - 1)
'        Next i
'    End With
'    Set myXinsFldrDelete = Nothing
'    Call variablesOfCFldrDelete(myXlonExeFldrCnt, myZstrExeFldrPath)     'Debug.Print
'JumpPath:
'End Sub
'Private Sub variablesOfCFldrDelete( _
'            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
''//CFldrDeleteクラス内から出力した変数の内容確認
'    Debug.Print "データ数: " & myXlonDataCnt
'    If myXlonDataCnt <= 0 Then Exit Sub
'  Dim k As Long
'    For k = LBound(myZvarField) To UBound(myZvarField)
'        Debug.Print "データ" & k & ": " & myZvarField(k)
'    Next k
'End Sub
