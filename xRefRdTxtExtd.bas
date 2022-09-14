Attribute VB_Name = "xRefRdTxtExtd"
'Includes CRdTxtNoOpn
'Includes CRdTxtNoOpnUTF8
'Includes CRdTxtOpn
'Includes CVrblToSht
'Includes CVrblToTxt
'Includes PfncstrGetTextFileCharset
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_テキストファイルの内容を取得してエクセルシートに書き出す
'◆ModuleProc名_テキストファイルの内容を取得してテキストファイルに書き出す
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRdTxtExtd"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数
  Private Const coXstrANSI As Variant = "Shift_JIS (ANSI)"
  Private Const coXstrUTF8 As Variant = "UTF-8"
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXstrFileCharset As String
    'myXstrFileCharset = Shift_JIS (ANSI)
    'myXstrFileCharset = UTF-8
    'myXstrFileCharset = UTF-8 BOM
    'myXstrFileCharset = UTF-16 LE BOM
    'myXstrFileCharset = UTF-16 BE BOM
    'myXstrFileCharset = EUC-JP
  Private myXstrDirPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String
  Private myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
            myZstrTxtData() As String
    'myZstrTxtData(i, j) : テキストファイル内容
    
'//入力制御信号
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : エクセルシートに書き出す
    'myXlonOutputOptn = 2 : テキストファイルに書き出す
  
'//入力データ
  Private myXstrOrgFilePath As String
  Private myXlonBgn As Long, myXlonEnd As Long, _
            myXbisSpltOptn As Boolean, myXstrInSpltChr As String
  Private myXobjPstFrstCell As Object
  Private myXstrPrntFilePath As String
  Private myXstrOutSpltChar As String
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myXbisMsgBoxON As Boolean
  Private myZvarPstData As Variant

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisInptBxOFF = False: myXbisEachWrtON = False
    myXbisMsgBoxON = False
    myZvarPstData = Empty
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
    Call callxRefRdTxtExtd
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXstrDirPathOUT As String, myXstrFileNameOUT As String, _
            myXstrBaseNameOUT As String, myXstrExtsnOUT As String, _
            myXlonTxtRowCntOUT As Long, myXlonTxtColCntOUT As Long, _
            myZstrTxtDataOUT() As String, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrOrgFilePathIN As String, _
            ByVal myXlonBgnIN As Long, ByVal myXlonEndIN As Long, _
            ByVal myXbisSpltOptnIN As Boolean, ByVal myXstrInSpltChrIN As String, _
            ByVal myXobjPstFrstCellIN As Object, _
            ByVal myXstrPrntFilePathIN As String, _
            ByVal myXstrOutSpltCharIN As String)

'//入力変数を初期化
    myXlonOutputOptn = Empty
    
    myXstrOrgFilePath = Empty
    myXlonBgn = Empty: myXlonEnd = Empty
    myXbisSpltOptn = False: myXstrInSpltChr = Empty
    Set myXobjPstFrstCell = Nothing
    myXstrPrntFilePath = Empty
    myXstrOutSpltChar = Empty

'//入力変数を取り込み
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrOrgFilePath = myXstrOrgFilePathIN
    myXlonBgn = myXlonBgnIN
    myXlonEnd = myXlonEndIN
    myXbisSpltOptn = myXbisSpltOptnIN
    myXstrInSpltChr = myXstrInSpltChrIN
    Set myXobjPstFrstCell = myXobjPstFrstCellIN
    myXstrPrntFilePath = myXstrPrntFilePathIN
    myXstrOutSpltChar = myXstrOutSpltCharIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
    myXstrDirPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    myXlonTxtRowCntOUT = Empty: myXlonTxtColCntOUT = Empty
    Erase myZstrTxtDataOUT

'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub

'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXstrDirPathOUT = myXstrDirPath
    myXstrFileNameOUT = myXstrFileName
    myXstrBaseNameOUT = myXstrBaseName
    myXstrExtsnOUT = myXstrExtsn
    myXlonTxtRowCntOUT = myXlonTxtRowCnt
    myXlonTxtColCntOUT = myXlonTxtColCnt
    If myXlonTxtRowCntOUT <= 0 Or myXlonTxtColCntOUT <= 0 Then Exit Sub
    myZstrTxtDataOUT() = myZstrTxtData()

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
    
'//S:テキストファイルの内容を取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:取得データを加工
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:変数情報を書き出す
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
    myXstrFileCharset = Empty
    myXstrDirPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    myXlonTxtRowCnt = Empty: myXlonTxtColCnt = Empty: Erase myZstrTxtData
End Sub

'RemP_モジュールメモリに保存した変数を取り出す
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
'    If myXlonOutputOptn < 0 And myXlonOutputOptn > 2 Then myXlonOutputOptn = 0
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_入力変数内容を確認する
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesA()
  
  Dim myXstrPrntPath As String, myXstrFileName As String
    myXstrPrntPath = ThisWorkbook.Path
    myXstrFileName = "testIN.txt"
    myXstrOrgFilePath = myXstrPrntPath & "\" & myXstrFileName
    
    myXlonBgn = 1
    myXlonEnd = 0
    
    myXbisSpltOptn = True
    myXstrInSpltChr = ""
    'myXbisSpltOptn = True  : 文字列を分割処理する
    'myXbisSpltOptn = False : 文字列を分割処理しない
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesB()
    
    myXlonOutputOptn = 1
    'myXlonOutputOptn = 0 : 書き出し処理無し
    'myXlonOutputOptn = 1 : エクセルシートに書き出す
    'myXlonOutputOptn = 2 : テキストファイルに書き出す
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables1()
    
    Set myXobjPstFrstCell = Selection
    
    myXbisInptBxOFF = False
    'myXbisInptBxOFF = False : 指定位置が無い場合にInputBoxで範囲指定する
    'myXbisInptBxOFF = True  : 指定位置が無い場合にInputBoxで範囲指定しない
    
    myXbisEachWrtON = False
    'myXbisEachWrtON = False : 配列変数内データを一度に書き出しする
    'myXbisEachWrtON = True  : 配列変数内データを1データづつ書き出しする
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
  Dim myXstrPrntPath As String, myXstrFileName As String
    myXstrPrntPath = ThisWorkbook.Path
    myXstrFileName = "testOUT.txt"
    myXstrPrntFilePath = myXstrPrntPath & "\" & myXstrFileName
    
    myXstrOutSpltChar = ""
    
    myXbisMsgBoxON = False
    'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
    'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
    
End Sub

'SnsP_
Private Sub snsProc()
    myXbisExitFlag = False
    
'//指定テキストファイルの文字コードを取得
    myXstrFileCharset = PfncstrGetTextFileCharset(myXstrOrgFilePath)
    If myXstrFileCharset = "" Then GoTo ExitPath
    
'//文字コードで処理を分岐
    Select Case myXstrFileCharset
        Case coXstrANSI
        '//ファイルを開かずにテキストファイルの内容を取得
            Call instCRdTxtNoOpn
            
        Case coXstrUTF8
        '//ファイルを開かずにUTF8形式テキストファイルの内容を取得
            Call instCRdTxtNoOpnUTF8
            
        Case Else
        '//ファイルを開いてテキストファイルの内容を取得
            Call instCRdTxtOpn
            
    End Select
    If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_取得データを加工
Private Sub prsProc()
    myXbisExitFlag = False
    
    On Error GoTo ExitPath
    myZvarPstData = myZstrTxtData
    On Error GoTo 0
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_
Private Sub runProc()
    myXbisExitFlag = False
    
'//変数情報を書き出す方法で分岐
    Select Case myXlonOutputOptn
    '//エクセルシートに書き出す
        Case 1
            Call setControlVariables1
            Call instCVrblToSht
        
    '//テキストファイルに書き出す
        Case 2
            Call setControlVariables2
            Call instCVrblToTxt
        
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

'◆ClassProc名_ファイルを開かずにテキストファイルの内容を取得する
Private Sub instCRdTxtNoOpn()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtNoOpn As CRdTxtNoOpn: Set myXinsRdTxtNoOpn = New CRdTxtNoOpn
    With myXinsRdTxtNoOpn
    '//クラス内変数への入力
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtNoOpn = Nothing
End Sub

'◆ClassProc名_ファイルを開かずにUTF8形式テキストファイルの内容を取得する
Private Sub instCRdTxtNoOpnUTF8()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtNoOpn As CRdTxtNoOpnUTF8: Set myXinsRdTxtNoOpn = New CRdTxtNoOpnUTF8
    With myXinsRdTxtNoOpn
    '//クラス内変数への入力
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtNoOpn = Nothing
End Sub

'◆ClassProc名_ファイルを開いてテキストファイルの内容を取得する
Private Sub instCRdTxtOpn()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtOpn As CRdTxtOpn: Set myXinsRdTxtOpn = New CRdTxtOpn
    With myXinsRdTxtOpn
    '//クラス内変数への入力
    '//テキストファイルパスを指定
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtOpn = Nothing
End Sub

'◆ClassProc名_変数情報をエクセルシートに書き出す
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//クラス内変数への入力
        .letVrbl = myZvarPstData
        Set .setPstFrstCell = myXobjPstFrstCell
        .letInptBxOFF = False
        .letEachWrtON = False
    '//クラス内プロシージャの実行とクラス内変数からの出力
        myXbisExitFlag = Not .fncbisCmpltFlag
    End With
    Set myXinsVrblToSht = Nothing
End Sub

'◆ClassProc名_変数情報をテキストファイルに書き出す
Private Sub instCVrblToTxt()
  Dim myXinsVrblToTxt As CVrblToTxt: Set myXinsVrblToTxt = New CVrblToTxt
    With myXinsVrblToTxt
    '//クラス内変数への入力
        .letVrbl = myZvarPstData
        .letSpltChar = myXstrOutSpltChar
        .letSaveFilePath = myXstrPrntFilePath
        .letMsgBoxON = False
    '//クラス内プロシージャの実行とクラス内変数からの出力
        myXbisExitFlag = Not .fncbisCmpltFlag
    End With
    Set myXinsVrblToTxt = Nothing
End Sub

'===============================================================================================
 
 '定型Ｆ_指定テキストファイルの文字コードを取得する
Private Function PfncstrGetTextFileCharset(ByVal myXstrFilePath As String) As String
'myXstrCharset = Shift_JIS (ANSI)
'myXstrCharset = UTF-8
'myXstrCharset = UTF-8 BOM
'myXstrCharset = UTF-16 LE BOM
'myXstrCharset = UTF-16 BE BOM
'myXstrCharset = EUC-JP
    PfncstrGetTextFileCharset = Empty
  Dim myXstrCharset As String, i As Long
  Dim myXlonHdlFile As Long, myXlonFileLen As Long
  Dim myZbytFile() As Byte, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
  Dim myXlonSJIS As Long, myXlonUTF8 As Long, myXlonEUC As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'//ファイル読み込み
    On Error Resume Next
    myXlonFileLen = FileLen(myXstrFilePath)
    ReDim myZbytFile(myXlonFileLen)
    If Err.Number <> 0 Then Exit Function
    myXlonHdlFile = FreeFile()
    Open myXstrFilePath For Binary As #myXlonHdlFile
    Get #myXlonHdlFile, , myZbytFile
    Close #myXlonHdlFile
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
'//BOMによる判断
    If (myZbytFile(L) = &HEF And myZbytFile(L + 1) = &HBB And myZbytFile(L + 2) = &HBF) Then
        myXstrCharset = "UTF-8 BOM"
        GoTo SetPath
    ElseIf (myZbytFile(L) = &HFF And myZbytFile(L + 1) = &HFE) Then
        myXstrCharset = "UTF-16 LE BOM"
        GoTo SetPath
    ElseIf (myZbytFile(L) = &HFE And myZbytFile(L + 1) = &HFF) Then
        myXstrCharset = "UTF-16 BE BOM"
        GoTo SetPath
    End If
'//BINARY
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If (b1 >= &H0 And b1 <= &H8) Or _
                (b1 >= &HA And b1 <= &H9) Or _
                (b1 >= &HB And b1 <= &HC) Or _
                (b1 >= &HE And b1 <= &H19) Or _
                (b1 >= &H1C And b1 <= &H1F) Or _
                (b1 = &H7F) Then
            myXstrCharset = "BINARY"
            GoTo SetPath
        End If
    Next i
'//Shift_JIS
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H9 Or b1 = &HA Or b1 = &HD Or _
                (b1 >= &H20 And b1 <= &H7E) Or _
                (b1 >= &HB0 And b1 <= &HDF) Then
            myXlonSJIS = myXlonSJIS + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                        ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   myXlonSJIS = myXlonSJIS + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
'//UTF-8
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H9 Or b1 = &HA Or b1 = &HD Or (b1 >= &H20 And b1 <= &H7E) Then
            myXlonUTF8 = myXlonUTF8 + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   myXlonUTF8 = myXlonUTF8 + 2
                   i = i + 1
                Else
                    If (i < myXlonFileLen - 3) Then
                        b3 = myZbytFile(i + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And _
                                (b2 >= &H80 And b2 <= &HBF) And _
                                (b3 >= &H80 And b3 <= &HBF) Then
                            myXlonUTF8 = myXlonUTF8 + 3
                            i = i + 2
                        Else
                            If (i < myXlonFileLen - 4) Then
                                b4 = myZbytFile(i + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And _
                                        (b2 >= &H80 And b2 <= &HBF) And _
                                        (b3 >= &H80 And b3 <= &HBF) And _
                                        (b4 >= &H80 And b4 <= &HBF) Then
                                    myXlonUTF8 = myXlonUTF8 + 4
                                    i = i + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
'//EUC-JP
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H7 Or b1 = 10 Or b1 = 13 Or (b1 >= &H20 And b1 <= &H7E) Then
            myXlonEUC = myXlonEUC + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And (b2 >= &HA1 And b2 <= &HFE)) Or _
                        (b1 = &H8E And (b2 >= &HA1 And b2 <= &HDF)) Then
                   myXlonEUC = myXlonEUC + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
'//文字コード出現順位による判断
    If (myXlonSJIS <= myXlonUTF8) And (myXlonEUC <= myXlonUTF8) Then
        myXstrCharset = "UTF-8"
        GoTo SetPath
    End If
    If (myXlonUTF8 <= myXlonSJIS) And (myXlonEUC <= myXlonSJIS) Then
        myXstrCharset = "Shift_JIS"
        GoTo SetPath
    End If
    If (myXlonUTF8 <= myXlonEUC) And (myXlonSJIS <= myXlonEUC) Then
        myXstrCharset = "EUC-JP"
        GoTo SetPath
    End If
    Exit Function
SetPath:
    PfncstrGetTextFileCharset = myXstrCharset
End Function

 '定型Ｐ_モジュール内定数の値を変更する
Private Sub PfixChangeModuleConstValue(myXbisExitFlag As Boolean, _
            ByVal myXstrMdlName As String, ByRef myZvarM() As Variant)
    myXbisExitFlag = False
    If myXstrMdlName = "" Then GoTo ExitPath
  Dim L As Long, myXvarTmp As Variant
    On Error GoTo ExitPath
    If IsArray(myZvarM) = False Then GoTo ExitPath
    L = LBound(myZvarM, 1): myXvarTmp = myZvarM(L, L)
    On Error GoTo 0
  Dim myXlonDclrLines As Long
    With ThisWorkbook.VBProject.VBComponents(myXstrMdlName).CodeModule
    myXlonDclrLines = .CountOfDeclarationLines
    If myXlonDclrLines < 1 Then GoTo ExitPath
  Dim i As Long, n As Long
  Dim myXstrTmp As String, myXstrRplcCode As String
    For i = 1 To myXlonDclrLines
        myXstrTmp = .Lines(i, 1)
        For n = LBound(myZvarM, 1) To UBound(myZvarM, 1)
          Dim myXstrSrch As String
            myXstrSrch = "Const" & Space(1) & myZvarM(n, L) & Space(1) & "As" & Space(1)
            If InStr(myXstrTmp, myXstrSrch) = 0 Then GoTo NextPath
          Dim myXstrOrg As String
            myXstrOrg = Left(myXstrTmp, InStr(myXstrTmp, "=" & Space(1)) + 1)
            myXstrRplcCode = myXstrOrg & myZvarM(n, L + 1)
            Application.DisplayAlerts = False
            Call .ReplaceLine(i, myXstrRplcCode)
            Application.DisplayAlerts = True
NextPath:
        Next n
    Next i
    End With
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'DummyＰ_
Private Sub MsubDummy()
End Sub

'===============================================================================================

''SetP_制御用変数を設定する
'Private Sub setControlVariables()
'    myXlonOutputOptn = 1
'    'myXlonOutputOptn = 1 : エクセルシートに書き出す
'    'myXlonOutputOptn = 2 : テキストファイルに書き出す
'  Dim myXstrPrntPath As String, myXstrFileName As String
'    myXstrPrntPath = ThisWorkbook.Path
'    myXstrFileName = "testIN.txt"
'    myXstrOrgFilePath = myXstrPrntPath & "\" & myXstrFileName
'    myXlonBgn = 1
'    myXlonEnd = 0
'    myXbisSpltOptn = True
'    myXstrInSpltChr = ""
'    'myXbisSpltOptn = True  : 文字列を分割処理する
'    'myXbisSpltOptn = False : 文字列を分割処理しない
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables1()
'    Set myXobjPstFrstCell = Selection
'End Sub
'
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'  Dim myXstrPrntPath As String, myXstrFileName As String
'    myXstrPrntPath = ThisWorkbook.Path
'    myXstrFileName = "testOUT.txt"
'    myXstrPrntFilePath = myXstrPrntPath & "\" & myXstrFileName
'    myXbisMsgBoxON = False
'    'myXbisMsgBxON = False : 変数のテキスト書き出し完了のMsgBoxを表示しない
'    'myXbisMsgBxON = True  : 変数のテキスト書き出し完了のMsgBoxを表示する
'    myXstrOutSpltChar = ""
'End Sub
'◆ModuleProc名_テキストファイルの内容を取得してエクセルシートに書き出す
'◆ModuleProc名_テキストファイルの内容を取得してテキストファイルに書き出す
Private Sub callxRefRdTxtExtd()
'  Dim myXlonOutputOptn As Long, _
'        myXstrOrgFilePath As String, myXlonBgn As Long, myXlonEnd As Long, _
'        myXbisSpltOptn As Boolean, myXstrInSpltChr As String, _
'        myXobjPstFrstCell As Object, _
'        myXstrPrntFilePath As String, myXstrOutSpltChar As String
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
'        myZstrTxtData() As String
    Call xRefRdTxtExtd.callProc( _
            myXbisCmpltFlag, _
            myXstrDirPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXlonTxtRowCnt, myXlonTxtColCnt, myZstrTxtData, _
            myXlonOutputOptn, _
            myXstrOrgFilePath, myXlonBgn, myXlonEnd, myXbisSpltOptn, myXstrInSpltChr, _
            myXobjPstFrstCell, myXstrPrntFilePath, myXstrOutSpltChar)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRdTxtExtd()
'//xRefRdTxtExtdモジュールのモジュールメモリのリセット処理
    Call xRefRdTxtExtd.resetConstant
End Sub
