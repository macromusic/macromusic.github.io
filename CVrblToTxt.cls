VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVrblToTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PincFileAutoNamingAndParentCheck
'Includes PfixGetFileNameInformation
'Includes PfncbisCheckFolderExist
'Includes PfncstrFileAutoNaming
'Includes PfnclonIsArrayExtd
'Includes PfnclonArrayDimension
'Includes PfixArrangeTextArrayVariant
'Includes PfixWriteTextLineWithoutFileOpening

Option Explicit
Option Base 1

'◆ClassProc名_変数情報をテキストファイルに書き出す
'Rev.008

'//モジュール内定数
  Private Const meMstrMdlName As String = "CVrblToTxt"
  
'//モジュール内定数
  Private Const coXstrTXT As String = "txt"
  Private Const coXstrCSV As String = "csv"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertySetLet_入力制御信号
  Private myXbisMsgBoxON As Boolean
  
'//PropertySetLet_入力データ
  Private myZvarVrbl As Variant, myXstrSpltChar As String
  Private myXstrSaveFilePath As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myZvarTxtSaveData As Variant, myXstrSaveTxtPath As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myZvarTxtSaveData = Empty: myXstrSaveTxtPath = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getCmpltFlag() As Boolean
    getCmpltFlag = myXbisCmpltFlag
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letVrbl(ByVal RHS As Variant)
    myZvarVrbl = RHS
End Property

Public Property Let letSpltChar(ByVal RHS As String)
    myXstrSpltChar = RHS
End Property

Public Property Let letSaveFilePath(ByVal RHS As String)
    myXstrSaveFilePath = RHS
End Property

Public Property Let letMsgBoxON(ByVal RHS As Boolean)
    myXbisMsgBoxON = RHS
'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'//ファイル名自動生成＆親フォルダチェック
  Dim myXbisFldrExistFlag As Boolean
    Call PincFileAutoNamingAndParentCheck( _
            myXbisFldrExistFlag, myXstrSaveTxtPath, _
            myXstrSaveFilePath, True)

'//入力変数が配列で空かどうかチェック
  Dim myXlonChck As Long: myXlonChck = PfnclonIsArrayExtd(myZvarVrbl)
    Select Case myXlonChck
        Case 0      '空の配列変数
            GoTo ExitPath
        Case 1      '空ではない配列変数
        Case -1     '配列ではない変数
    End Select
    
'//配列変数の次元数を取得
  Dim myXlonAryDmnsn As Long: myXlonAryDmnsn = PfnclonArrayDimension(myZvarVrbl)
    
'//配列変数の次元数に応じて処理を実施
  Dim i As Long, Li As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXlonRowCnt As Long, myXlonColCnt As Long
    If myXlonAryDmnsn = 0 Then
        If CStr(myZvarVrbl) = "" Then GoTo ExitPath
        ReDim myZvarTxtSaveData(Lo, Lo) As Variant
        myZvarTxtSaveData(Lo, Lo) = myZvarVrbl
        
    ElseIf myXlonAryDmnsn = 1 Then
        Li = LBound(myZvarVrbl)
        i = UBound(myZvarVrbl) + Lo - Li
        ReDim myZvarTxtSaveData(i, Lo) As Variant
        For i = LBound(myZvarVrbl) To UBound(myZvarVrbl)
            myZvarTxtSaveData(i + Lo - Li, Lo) = myZvarVrbl(i)
        Next i
        
    ElseIf myXlonAryDmnsn = 2 Then
        If UBound(myZvarVrbl, 2) = LBound(myZvarVrbl, 2) Then
            myZvarTxtSaveData = myZvarVrbl
            
        Else
        '//テキストファイルに書き込む時の結合文字の設定
            Call checkInputVariablesSub
            If myXbisExitFlag = True Then GoTo ExitPath
            
        '//テキストファイル形式の配列に入れ替え
            Call PfixArrangeTextArrayVariant( _
                    myXlonRowCnt, myXlonColCnt, myZvarTxtSaveData, _
                    myZvarVrbl, myXstrSpltChar)
            
        End If
        
    Else
        GoTo ExitPath
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariablesSub()
    myXbisExitFlag = False
    
'//テキストファイルに書き込む時の結合文字の設定
    If myXstrSpltChar <> "" Then Exit Sub
    
'//指定ファイルのファイル名情報を取得
  Dim myXstrPrntPath As String, myXstrFileName As String, _
        myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformation( _
            myXstrPrntPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrSaveFilePath)
    
'//結合処理する文字列を設定
'Chr(9)  : タブ文字(vbTab)
'Chr(44) : カンマ(,)
    Select Case LCase(myXstrExtsn)
        Case LCase(coXstrTXT): myXstrSpltChar = vbTab
        Case LCase(coXstrCSV): myXstrSpltChar = Chr(44)
        Case Else: GoTo ExitPath
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

'PublicF_
Public Function fncbisCmpltFlag() As Boolean
    fncbisCmpltFlag = Empty
    Call ctrProc
    fncbisCmpltFlag = myXbisCmpltFlag
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    
    Call checkInputVariables
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"  'PassFlag
    
'//変数情報をテキストファイルに書出し
    Call PfixWriteTextLineWithoutFileOpening( _
            myXbisExitFlag, myZvarTxtSaveData, myXstrSaveTxtPath, False)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
    If myXbisMsgBoxON = True Then
        MsgBox "変数のテキスト書き出し完了。" _
                & vbCrLf & vbCrLf & myXstrSaveTxtPath _
                & vbCrLf & vbCrLf & "上記ファイルを確認下さい。"
    End If

    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

'===============================================================================================

 '定型Ｐ_ファイル名自動生成＆親フォルダチェック
Private Sub PincFileAutoNamingAndParentCheck( _
            myXbisFldrExistFlag As Boolean, myXstrFilePathOUT As String, _
            ByVal myXstrFilePathIN As String, _
            Optional ByVal coXbisAutoNaming As Boolean = False)
'Includes PfixGetFileNameInformation
'Includes PfncbisCheckFolderExist
'Includes PfncstrFileAutoNaming
'coXbisAutoNaming = False  : 同一名ファイルがあれば終了
'coXbisAutoNaming = True   : 同一名ファイルがあればファイル名を自動生成
    myXbisFldrExistFlag = False: myXstrFilePathOUT = Empty
'//指定ファイルのファイル名情報を取得
  Dim myXstrPrntPath As String, myXstrFileName As String, _
        myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformation( _
            myXstrPrntPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrFilePathIN)
    If myXstrBaseName = "" Or myXstrExtsn = "" Then GoTo ExitPath
'//指定フォルダの存在を確認
    myXbisFldrExistFlag = PfncbisCheckFolderExist(myXstrPrntPath)
    If myXbisFldrExistFlag = False Then myXstrPrntPath = ThisWorkbook.Path
'//ファイル名を自動生成
    myXstrFilePathOUT = PfncstrFileAutoNaming( _
                            myXstrPrntPath, myXstrBaseName, myXstrExtsn)
    Exit Sub
ExitPath:
    myXstrFilePathOUT = ""
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

 '定型Ｆ_指定フォルダの存在を確認する
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

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

 '定型Ｆ_指定変数が配列で空かどうかチェックする
Private Function PfnclonIsArrayExtd(ByRef myZvarOrgData As Variant) As Long
'PfnclonIsArrayExtd = 1  : 空ではない配列変数
'PfnclonIsArrayExtd = 0  : 空の配列変数
'PfnclonIsArrayExtd = -1 : 配列ではない変数
    PfnclonIsArrayExtd = Empty
  Dim myXlonChckFlag As Long, myXlonTmp As Long
'//配列であることを確認
    If IsArray(myZvarOrgData) = False Then
        myXlonChckFlag = -1: GoTo JumpPath
    End If
'//配列が空でないことを確認
    myXlonChckFlag = 0
    On Error GoTo JumpPath
    myXlonTmp = UBound(myZvarOrgData) - LBound(myZvarOrgData) + 1
    On Error GoTo 0
    If myXlonTmp <= 0 Then GoTo JumpPath
    myXlonChckFlag = 1
JumpPath:
    PfnclonIsArrayExtd = myXlonChckFlag
End Function

 '定型Ｆ_配列変数の次元数を取得する
Private Function PfnclonArrayDimension(ByRef myZvarOrgData As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarOrgData) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
End Function

 '定型Ｐ_2次元配列のVariant変数をテキストファイル形式の配列に入れ替える
Private Sub PfixArrangeTextArrayVariant( _
            myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
            myZvarTxtSaveData As Variant, _
            ByRef myZvarOrgData As Variant, _
            Optional ByVal coXstrSpltChar As String = vbTab)
'myZvarTxtSaveData(i, 1) : 入れ替え後データ配列
'myZvarOrgData(i, j) : 元データ配列
    myXlonTxtRowCnt = Empty: myXlonTxtColCnt = Empty: myZvarTxtSaveData = Empty
    If coXstrSpltChar = "" Then Exit Sub
  Dim myXvarTmp As Variant, Li As Long, Ui1 As Long
    On Error GoTo ExitPath
    Li = LBound(myZvarOrgData, 1): myXvarTmp = myZvarOrgData(Li, Li)
    Ui1 = UBound(myZvarOrgData, 1)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
    myXlonTxtRowCnt = Ui1 - Li + 1: myXlonTxtColCnt = 1
    i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
    ReDim myZvarTxtSaveData(i, j) As Variant
    For i = Li To Ui1
        myXvarTmp = Empty
        For j = LBound(myZvarOrgData, 2) To UBound(myZvarOrgData, 2)
            If j = LBound(myZvarOrgData, 2) Then
                myXvarTmp = myZvarOrgData(i, j)
            Else
                myXvarTmp = myXvarTmp & coXstrSpltChar & myZvarOrgData(i, j)
            End If
            myZvarTxtSaveData(i + Lo - Li, Lo) = myXvarTmp
        Next j
    Next i
ExitPath:
End Sub

 '定型Ｐ_ファイルを開かずにテキストファイルに1行づつテキストデータを書き込む
Private Sub PfixWriteTextLineWithoutFileOpening( _
            myXbisExitFlag As Boolean, _
            ByRef myZvarTxtData As Variant, ByVal myXstrTxtFilePath As String, _
            Optional ByVal coXbisAppend As Boolean = False)
'myZvarTxtData(i, 1) : 元データ配列
'coXbisAppend = False : 上書きモードでテキストファイルに書き込む
'coXbisAppend = True  : 追加モードでテキストファイルに書き込む
    myXbisExitFlag = False
    If myXstrTxtFilePath = "" Then GoTo ExitPath
'//入力変数の次元数を確認
  Dim myXlonTmp As Long, n As Long: n = 0
    On Error Resume Next
    Do
        n = n + 1: myXlonTmp = UBound(myZvarTxtData, n)
    Loop While Err.Number = 0
    On Error GoTo 0
    If n - 1 <> 2 Then GoTo ExitPath
'//入力変数をテキストファイルに書き込み
  Dim myXlonFNo As Long: myXlonFNo = FreeFile
    If coXbisAppend = False Then
        Open myXstrTxtFilePath For Output As #myXlonFNo
    Else
        Open myXstrTxtFilePath For Append As #myXlonFNo
    End If
  Dim i As Long
    For i = LBound(myZvarTxtData, 1) To UBound(myZvarTxtData, 1)
        Print #myXlonFNo, myZvarTxtData(i, 1)
    Next i
    Close #myXlonFNo
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'DummyＰ_
Private Sub CsubDummy()
End Sub

'===============================================================================================

''◆ClassProc名_変数情報をテキストファイルに書き出す
'Private Sub instCVrblToTxt()
'  Dim myXbisCmpltFlag As Boolean
''//クラス内変数への入力
'  Dim myXbisMsgBoxON As Boolean
'    myXbisMsgBoxON = False
'    'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
'    'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
'  Dim myZvarVrbl As Variant
'    ReDim myZvarVrbl(2, 2) As Variant
'    myZvarVrbl(1, 1) = "A"
'    myZvarVrbl(1, 2) = "A"
'    myZvarVrbl(2, 1) = "A"
'    myZvarVrbl(2, 2) = "A"
'  Dim myXstrSpltChar As String
'    myXstrSpltChar = ""
'  Dim myXstrSaveFilePath As String
'    myXstrSaveFilePath = ""
''//CVrblToTxtクラスモジュールのインスタンス生成
'  Dim myXinsVrblToTxt As CVrblToTxt: Set myXinsVrblToTxt = New CVrblToTxt
'    With myXinsVrblToTxt
'    '//クラス内変数への入力
'        .letVrbl = myZvarVrbl
'        .letSpltChar = myXstrSpltChar
'        .letSaveFilePath = myXstrSaveFilePath
'        .letMsgBoxON = True
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXbisCmpltFlag = .getCmpltFlag
''        myXbisCmpltFlag = .fncbisCmpltFlag
'    End With
'    Set myXinsVrblToTxt = Nothing
''    Debug.Print "完了: " & myXbisCmpltFlag
'End Sub
