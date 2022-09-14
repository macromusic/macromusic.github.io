Attribute VB_Name = "xRefSubFldrLstExtd"
'Includes CSubFldrLst
'Includes CVrblToSht
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_指定ディレクトリ内のサブフォルダ一覧を取得してシートに書き出す
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSubFldrLstExtd"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXlonFldrCnt As Long, myZobjFldr() As Object, _
            myZstrFldrName() As String, myZstrFldrPath() As String
    'myZobjFldr(k) : フォルダオブジェクト
    'myZstrFldrName(k) : フォルダ名
    'myZstrFldrPath(k) : フォルダパス
  Private myXobjPstdCell As Object
  
'//入力制御信号
  Private myXbisNotOutFldrInfo As Boolean
    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
  
'//入力データ
  Private myXstrDirPath As String
  
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す
  
  Private myXobjFldrPstFrstCell As Object
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXobjDir As Object

  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myZvarPstVrbl As Variant
    
'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjDir = Nothing
    myXbisInptBxOFF = False: myXbisEachWrtON = False
    myZvarPstVrbl = Empty
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_モジュールメモリのリセット
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()
    
'//処理実行
    Call callxRefSubFldrLstExtd
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            myXobjPstdCellOUT As Object, _
            ByVal myXbisNotOutFldrInfoIN As Boolean, _
            ByVal myXstrDirPathIN As String, _
            ByVal myXlonOutputOptnIN As Long, ByVal myXobjFldrPstFrstCellIN As Object)
    
'//入力変数を初期化
    myXbisNotOutFldrInfo = False
    myXstrDirPath = Empty
    myXlonOutputOptn = Empty
    Set myXobjFldrPstFrstCell = Nothing

'//入力変数を取り込み
    myXbisNotOutFldrInfo = myXbisNotOutFldrInfoIN
    myXstrDirPath = myXstrDirPathIN
    myXlonOutputOptn = myXlonOutputOptnIN
    Set myXobjFldrPstFrstCell = myXobjFldrPstFrstCellIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT: Erase myZstrFldrNameOUT: Erase myZstrFldrPathOUT
    Set myXobjPstdCellOUT = Nothing
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    Set myXobjPstdCellOUT = myXobjPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariablesA
    Call setControlVariablesB
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:指定ディレクトリ内のサブフォルダ一覧を取得
    Call instCSubFldrLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:取得データを加工
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:変数情報をエクセルシートに書き出す
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonFldrCnt = Empty
    Erase myZobjFldr: Erase myZstrFldrName: Erase myZstrFldrPath
End Sub

'RemP_モジュールメモリに保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'    If myXstrDirPath = "" Then GoTo ExitPath
'
'    If myXlonOutputOptn < 0 Or myXlonOutputOptn > 2 Then myXlonDirSlctOptn = 1
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesA()

    myXbisNotOutFldrInfo = False
    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
    
'    myXstrDirPath = ActiveWorkbook.Path
    myXstrDirPath = "C:\Users\Hiroki\Documents\_VBA4XPC\11 プログラムデータベース\02_VBAモジュール"

End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesB()
    
    myXlonOutputOptn = 1
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す

'    myZvarVrbl = 1
    
'    Set myXobjFldrPstFrstCell = Selection
    
    myXbisInptBxOFF = False
    'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
    'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
    
    myXbisEachWrtON = False
    'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
    'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
    
End Sub

'SnsP_
Private Sub snsProc()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_取得データを加工
Private Sub prsProc()
    myXbisExitFlag = False
    
    On Error GoTo ExitPath
    Select Case myXlonOutputOptn
    '//フォルダパスを選択
        Case 1: myZvarPstVrbl = myZstrFldrPath
        
    '//フォルダ名を選択
        Case 2: myZvarPstVrbl = myZstrFldrName
        
        Case Else: Exit Sub
    End Select
    On Error GoTo 0
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_変数情報をエクセルシートに書き出す
Private Sub runProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt1 As String _
        = "フォルダパスを貼り付ける位置を指定して下さい。"
  Const coXstrMsgBxPrmpt2 As String _
        = "フォルダ名を貼り付ける位置を指定して下さい。"
    
'//変数情報を書き出すかで分岐
    Select Case myXlonOutputOptn
    '//エクセルシートに書き出す
        Case 1
            If myXbisInptBxOFF = False And myXobjFldrPstFrstCell Is Nothing Then _
                MsgBox coXstrMsgBxPrmpt1
            Call instCVrblToSht
        
    '//エクセルシートに書き出す
        Case 2
            If myXbisInptBxOFF = False And myXobjFldrPstFrstCell Is Nothing Then _
                MsgBox coXstrMsgBxPrmpt2
            Call instCVrblToSht
        
        Case Else: Exit Sub
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_出力変数内容を確認する
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_使用した変数をモジュールメモリに保存する
Private Sub recProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum"
    myZvarM(1, 2) = meMlonExeNum + 1

  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

'◆ClassProc名_指定ディレクトリ内のサブフォルダ一覧を取得する
Private Sub instCSubFldrLst()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSubFldrLst As CSubFldrLst: Set myXinsSubFldrLst = New CSubFldrLst
    With myXinsSubFldrLst
    '//クラス内変数への入力
        .letNotOutFldrInfo = myXbisNotOutFldrInfo
        .letDirPath = myXstrDirPath
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonFldrCnt = .getFldrCnt
        If myXlonFldrCnt <= 0 Then GoTo JumpPath
        k = myXlonFldrCnt + Lo - 1
        ReDim myZobjFldr(k) As Object
        ReDim myZstrFldrName(k) As String
        ReDim myZstrFldrPath(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFldrCnt
            Set myZobjFldr(k + Lo - 1) = .getFldrAry(k + Lc - 1)
            myZstrFldrName(k + Lo - 1) = .getFldrNameAry(k + Lc - 1)
            myZstrFldrPath(k + Lo - 1) = .getFldrPathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSubFldrLst = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSubFldrLst = Nothing
End Sub

'◆ClassProc名_変数情報をエクセルシートに書き出す
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//クラス内変数への入力
        .letVrbl = myZvarPstVrbl
        Set .setPstFrstCell = myXobjFldrPstFrstCell
        .letInptBxOFF = myXbisInptBxOFF
        .letEachWrtON = myXbisEachWrtON
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXbisExitFlag = Not .getCmpltFlag
        Set myXobjPstdCell = .getPstdRng
    End With
    Set myXinsVrblToSht = Nothing
End Sub

'===============================================================================================

 '定型Ｐ_モジュール内定数の値を変更する
Private Sub PfixChangeModuleConstValue(myXbisExitFlag As Boolean, _
            ByVal myXstrMdlName As String, ByRef myZvarM() As Variant)
    myXbisExitFlag = False
    If myXstrMdlName = "" Then GoTo ExitPath
  Dim L As Long, myXvarTmp As Variant
    On Error GoTo ExitPath
    L = LBound(myZvarM, 1): myXvarTmp = myZvarM(L, L)
    On Error GoTo 0
  Dim myXlonDclrLines As Long, myXobjCdMdl As Object
    Set myXobjCdMdl = ThisWorkbook.VBProject.VBComponents(myXstrMdlName).CodeModule
    myXlonDclrLines = myXobjCdMdl.CountOfDeclarationLines
    If myXlonDclrLines <= 0 Then GoTo ExitPath
  Dim i As Long, n As Long
  Dim myXstrTmp As String, myXstrSrch As String, myXstrOrg As String, myXstrRplc As String
Application.DisplayAlerts = False
    For i = 1 To myXlonDclrLines
        myXstrTmp = myXobjCdMdl.Lines(i, 1)
        For n = LBound(myZvarM, 1) To UBound(myZvarM, 1)
            myXstrSrch = "Const" & Space(1) & myZvarM(n, L) & Space(1) & "As" & Space(1)
            If InStr(myXstrTmp, myXstrSrch) > 0 Then
                myXstrOrg = Left(myXstrTmp, InStr(myXstrTmp, "=" & Space(1)) + 1)
                myXstrRplc = myXstrOrg & myZvarM(n, L + 1)
                Call myXobjCdMdl.ReplaceLine(i, myXstrRplc)
            End If
        Next n
    Next i
Application.DisplayAlerts = True
    Set myXobjCdMdl = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'DummyＰ_
Private Sub MsubDummy()
End Sub

'===============================================================================================

''SetP_制御用変数を設定する
'Private Sub setControlVariablesA()
'    myXbisNotOutFldrInfo = False
'    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
'    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
'    myXstrDirPath = ActiveWorkbook.Path
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariablesB()
'    myXlonOutputOptn = 1
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : フォルダパスをエクセルシートに書き出す
'    'myXlonOutputOptn = 2 : フォルダ名をエクセルシートに書き出す
''    myZvarVrbl = 1
''    Set myXobjFldrPstFrstCell = Selection
'End Sub
'◆ModuleProc名_指定ディレクトリ内のサブフォルダ一覧を取得してシートに書き出す
Private Sub callxRefSubFldrLstExtd()
'  Dim myXbisNotOutFldrInfo As Boolean, myXstrDirPath As String, _
'        myXlonOutputOptn As Long, myXobjFldrPstFrstCell As Object
'    'myXbisNotOutFldrInfo = False : フォルダオブジェクとフォルダ情報を両方出力する
'    'myXbisNotOutFldrInfo = True  : フォルダオブジェクトのみ出力してフォルダ情報は出力しない
'    'myXlonOutputOptn = 0 : 書き出し処理無し
'    'myXlonOutputOptn = 1 : エクセルシートに書き出す
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String, myXobjPstdCell As Object
'    'myZobjFldr(k) : フォルダオブジェクト
'    'myZstrFldrName(k) : フォルダ名
'    'myZstrFldrPath(k) : フォルダパス
    Call xRefSubFldrLstExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, myXobjPstdCell, _
            myXbisNotOutFldrInfo, myXstrDirPath, myXlonOutputOptn, myXobjFldrPstFrstCell)
    Debug.Print "データ: " & myXlonFldrCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrLstExtd()
'//xRefSubFldrLstExtdモジュールのモジュールメモリのリセット処理
    Call xRefSubFldrLstExtd.resetConstant
End Sub
