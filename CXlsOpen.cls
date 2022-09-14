VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CXlsOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Includes PfnclonCheckExcelBookOpening
'Includes PfixGetFileNameInformationByFSO
'Includes PfixOpenExcelBookDesignateOpeningMethod
'Includes PfncbisCheckExcelBookReadOnly
'Includes PfncobjGetExcelBookIfOpened
'Includes PfixCloseExcelBook

Option Explicit
Option Base 1

'◆ClassProc名_エクセルブックを開く
'Rev.003
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "CXlsOpen"
  
'//モジュール内定数
  Private Const coXstrXL As String = "xl"
  Private Const coXstrCSV As String = "csv"
  Private Const coXstrTXT As String = "txt"
  Private Const coXstrLVM As String = "lvm"
  
'//PropertyGet_出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//PropertyGet_出力データ
  Private myXobjOpndBook As Object
  
'//PropertySetLet_入力制御信号
  Private myXbisOpnRdOnly As Boolean
  
'//PropertySetLet_入力データ
  Private myXstrOpnFullName As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  Private myXlonBkOpndType As Long, myXbisBkRdOnly As Boolean
  
'//モジュール内変数_データ
  Private myXstrOpnBkName As String, myXlonExtsnFrmt As Long

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonBkOpndType = Empty: myXbisBkRdOnly = False
    
    myXstrOpnBkName = Empty: myXlonExtsnFrmt = Empty
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Get getCmpltFlag() As Boolean
    getCmpltFlag = myXbisCmpltFlag
End Property

Public Property Get getOpndBook() As Object
    Set getOpndBook = myXobjOpndBook
End Property

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    Set myXobjOpndBook = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

Public Property Let letOpnRdOnly(ByVal RHS As Boolean)
    myXbisOpnRdOnly = RHS
'myXbisOpnRdOnly = False : 指定エクセルブックを読み取り専用にせずに開く
'myXbisOpnRdOnly = True  : 指定エクセルブックを読み取り専用で開く
End Property

Public Property Let letOpnFullName(ByVal RHS As String)
    myXstrOpnFullName = RHS
End Property

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'//指定エクセルブックが既に開いているか確認
    myXlonBkOpndType = PfnclonCheckExcelBookOpening(myXstrOpnFullName)
    If myXlonBkOpndType <= 0 Then
        GoTo ExitPath
    ElseIf myXlonBkOpndType >= 4 Then
        GoTo ExitPath
    End If
    
'//指定ファイルのファイル名情報を取得
  Dim myXstrPrntPath As String, myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformationByFSO( _
            myXstrPrntPath, myXstrOpnBkName, myXstrBaseName, myXstrExtsn, _
            myXstrOpnFullName)
  Dim myXstrTmp As String, myXbisExtnFlag As Boolean: myXbisExtnFlag = False
    myXstrTmp = Left(myXstrExtsn, Len(coXstrXL))
    If LCase(myXstrTmp) = LCase(coXstrXL) Then
        myXbisExtnFlag = True
        myXlonExtsnFrmt = 5
    End If
    
    myXstrTmp = Left(myXstrExtsn, Len(coXstrCSV))
    If LCase(myXstrTmp) = LCase(coXstrCSV) Then
        myXbisExtnFlag = True
        myXlonExtsnFrmt = 2
    End If
    
    myXstrTmp = Left(myXstrExtsn, Len(coXstrTXT))
    If LCase(myXstrTmp) = LCase(coXstrTXT) Then
        myXbisExtnFlag = True
        myXlonExtsnFrmt = 1
    End If
    
    myXstrTmp = Left(myXstrExtsn, Len(coXstrLVM))
    If LCase(myXstrTmp) = LCase(coXstrLVM) Then
        myXbisExtnFlag = True
        myXlonExtsnFrmt = 2
    End If
    
    If myXbisExtnFlag = False Then GoTo ExitPath
    
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
Public Function fncobjOpenedBook() As Object
    Set fncobjOpenedBook = Nothing
    Call ctrProc
    Set fncobjOpenedBook = myXobjOpndBook
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag

    If myXlonBkOpndType = 1 Then
    '//指定エクセルブックが開いていない場合の処理
        Call runProc1
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
        
    ElseIf myXlonBkOpndType = 2 Then
    '//指定エクセルブックと同一名の別ブックが開いている場合の処理
        Call runProc2
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
        
    ElseIf myXlonBkOpndType = 3 Then
    '//指定エクセルブックが開いている場合の処理
        Call runProc3
        If myXbisExitFlag = True Then GoTo ExitPath
    '    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
        
    End If
    
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

'RunP_指定エクセルブックが開いていない場合の処理
Private Sub runProc1()
    myXbisExitFlag = False
        
'//指定エクセルブックを開く
    If myXlonExtsnFrmt < 1 Or myXlonExtsnFrmt > 6 Then myXlonExtsnFrmt = 5
    Set myXobjOpndBook = PfncobjOpenExcelBookByOpeningMethod( _
                            myXstrOpnFullName, _
                            0, myXbisOpnRdOnly, myXlonExtsnFrmt, "", True)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_指定エクセルブックと同一名の別ブックが開いている場合の処理
Private Sub runProc2()
    myXbisExitFlag = False
        
'//エクセルブックを閉じる
    Call PfixCloseExcelBook(myXbisExitFlag, myXstrOpnBkName)
    If myXbisExitFlag = True Then GoTo ExitPath
    
'//指定エクセルブックを開く
    If myXlonExtsnFrmt < 1 Or myXlonExtsnFrmt > 6 Then myXlonExtsnFrmt = 5
    Set myXobjOpndBook = PfncobjOpenExcelBookByOpeningMethod( _
                            myXstrOpnFullName, _
                            0, myXbisOpnRdOnly, myXlonExtsnFrmt, "", True)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_指定エクセルブックが開いている場合の処理
Private Sub runProc3()
    myXbisExitFlag = False
    
'//開いたエクセルブックが読み取り専用か確認
    myXbisBkRdOnly = PfncbisCheckExcelBookReadOnly(myXstrOpnBkName)
    
    If myXbisBkRdOnly = myXbisOpnRdOnly Then
    '//指定名のエクセルブックが既に開いていればブックオブジェクトを取得
        Set myXobjOpndBook = PfncobjGetExcelBookIfOpened(myXstrOpnBkName)
        
    Else
    '//エクセルブックを閉じる
        Call PfixCloseExcelBook(myXbisExitFlag, myXstrOpnBkName)
        If myXbisExitFlag = True Then GoTo ExitPath
        
    '//指定エクセルブックを開く
        If myXlonExtsnFrmt < 1 Or myXlonExtsnFrmt > 6 Then myXlonExtsnFrmt = 5
        Set myXobjOpndBook = PfncobjOpenExcelBookByOpeningMethod( _
                                myXstrOpnFullName, _
                                0, myXbisOpnRdOnly, myXlonExtsnFrmt, "", True)
        
    End If
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
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
    myXstrBookName = myXobjFSO.GetFileName(myXstrFullName)
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
        myXstrFileName = .GetFileName(myXstrFilePath)           'ファイル名
        myXstrBaseName = .GetBaseName(myXstrFilePath)           'ファイルベース名
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         'ファイル拡張子
    End With
    Set myXobjFSO = Nothing
End Sub

 '定型Ｆ_エクセルブックを開き方を指定して開く
Private Function PfncobjOpenExcelBookByOpeningMethod( _
            ByVal myXstrFullName As String, _
            Optional ByVal coXlonUpdateLinks As Long = 0, _
            Optional ByVal coXbisReadOnly As Boolean = False, _
            Optional ByVal coXlonFormat As Long = 1, _
            Optional ByVal coXstrDelimiter As String = "", _
            Optional ByVal coXbisIgnoreReadOnlyRecommended As Boolean = False) As Object
'UpdateLinks:=0 : リンク更新しない
'UpdateLinks:=3 : リンク更新する
'ReadOnly:=False : 通常
'ReadOnly:=True  : 読取り専用
'Format:=1 : ダブ区切り          : テキスト形式ファイルの開き方
'Format:=2 : コンマ(,)区切り     : テキスト形式ファイルの開き方
'Format:=3 : スペース区切り      : テキスト形式ファイルの開き方
'Format:=4 : セミコロン(;)区切り : テキスト形式ファイルの開き方
'Format:=5 : 区切り指定無し      : テキスト形式ファイルの開き方
'Format:=6 : カスタム文字区切り(引数Delimiterで指定) : テキスト形式ファイルの開き方
'IgnoreReadOnlyRecommended:=False : [読み取り専用を推奨する] メッセージを表示する
'IgnoreReadOnlyRecommended:=True  : [読み取り専用を推奨する] メッセージを表示しない
    Set PfncobjOpenExcelBookByOpeningMethod = Nothing
    On Error Resume Next
    Set PfncobjOpenExcelBookByOpeningMethod _
            = Workbooks.Open(Filename:=myXstrFullName, _
                        UpdateLinks:=coXlonUpdateLinks, _
                        ReadOnly:=coXbisReadOnly, _
                        Format:=coXlonFormat, _
                        Delimiter:=coXstrDelimiter, _
                        IgnoreReadOnlyRecommended:=coXbisIgnoreReadOnlyRecommended)
    On Error GoTo 0
End Function

 '定型Ｆ_開いたエクセルブックが読み取り専用か確認する
Private Function PfncbisCheckExcelBookReadOnly( _
            ByVal myXstrBookName As String) As Boolean
'PfncbisCheckExcelBookReadOnly = True  : 読み取り専用
'PfncbisCheckExcelBookReadOnly = False : 読み取り専用ではない
    PfncbisCheckExcelBookReadOnly = False
    On Error GoTo ExitPath
  Dim myXobjBook  As Object: Set myXobjBook = Workbooks(myXstrBookName)
    On Error GoTo 0
    PfncbisCheckExcelBookReadOnly = myXobjBook.ReadOnly
    Set myXobjBook = Nothing
    Exit Function
ExitPath:
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

 '定型Ｐ_エクセルブックを閉じる
Private Sub PfixCloseExcelBook(myXbisExitFlag As Boolean, _
            ByVal myXstrBookName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Workbooks(myXstrBookName).Close SaveChanges:=False
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

''◆ClassProc名_エクセルブックを開く
'Private Sub instCXlsOpen()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjOpndBook As Object
''//クラス内変数への入力
'  Dim myXbisOpnRdOnly As Boolean: myXbisOpnRdOnly = False
'    'myXbisOpnRdOnly = False : 指定エクセルブックを読み取り専用にせずに開く
'    'myXbisOpnRdOnly = True  : 指定エクセルブックを読み取り専用で開く
'  Dim myXstrOpnFullName As String
'    myXstrOpnFullName = ""
''//CXlsOpenクラスモジュールのインスタンス生成
'  Dim myXinsXlsOpen As CXlsOpen: Set myXinsXlsOpen = New CXlsOpen
'    With myXinsXlsOpen
'    '//クラス内変数への入力
'        .letOpnRdOnly = myXbisOpnRdOnly
'        .letOpnFullName = myXstrOpnFullName
'    '//クラス内プロシージャの実行とクラス内変数からの出力
'        .exeProc
'        myXbisCmpltFlag = .getCmpltFlag
'        Set myXobjOpndBook = .getOpndBook
'        Set myXobjOpndBook = .fncobjOpenedBook
'    End With
'    Set myXinsXlsOpen = Nothing
''    Debug.Print "データ: " & myXobjOpndBook.Name
'End Sub
