Attribute VB_Name = "xRefSlctShtData"
'Includes CSlctShtSrsData
'Includes CSlctShtDscrtData
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上の範囲を指定してその範囲のデータと情報を取得する
'Rev.002
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefSlctShtData"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
  
  Private myXlonSrsDataRowCnt As Long, myXlonSrsDataColCnt As Long, _
            myZstrShtSrsData() As String, myZvarShtSrsData() As Variant, _
            myZstrCmntData() As String
    'myZstrShtSrsData(i, j) : 取得文字列
    'myZvarShtSrsData(i, j) : 取得文字列
    'myZstrCmntData(i, j) : 取得コメント
  Private myXlonBgnRow As Long, myXlonEndRow As Long, _
            myXlonBgnCol As Long, myXlonEndCol As Long, _
            myXlonRows As Long, myXlonCols As Long
  
  Private myXlonDscrtDataCnt As Long, myZobjDscrtDataCell() As Object, _
            myZvarShtDscrtData() As Variant
    'myZvarShtDscrtData(i, 1) = Row
    'myZvarShtDscrtData(i, 2) = Column
    'myZvarShtDscrtData(i, 3) = SheetData
    'myZvarShtDscrtData(i, 4) = CommentData
    
'//入力制御信号
  Private myXlonSrsDataOptn As Long
    'myXlonSrsDataOptn = 1 : 連続データを取得する
    'myXlonSrsDataOptn = 2 : 不連続データを取得する
  
'//入力データ
  Private myXlonRngOptn As Long
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
  Private myXbisByVrnt As Boolean
    'myXbisByVrnt = False : シートデータをStringで取得する
    'myXbisByVrnt = True  : シートデータをVariantで取得する
  Private myXbisGetCmnt As Boolean
    'myXbisGetCmnt = False : コメントを取得しない
    'myXbisGetCmnt = True  : コメントを取得する
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean

'//モジュール内変数_データ
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrInptBxPrmpt = Empty: myXstrInptBxTtl = Empty
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
    Call callxRefSlctShtData
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXobjBookOUT As Object, myXstrShtNameOUT As String, myXlonShtNoOUT As Long, _
            myXlonSrsDataRowCntOUT As Long, myXlonSrsDataColCntOUT As Long, _
            myZstrShtSrsDataOUT() As String, myZvarShtSrsDataOUT() As Variant, _
            myZstrCmntDataOUT() As String, _
            myXlonBgnRowOUT As Long, myXlonEndRowOUT As Long, _
            myXlonBgnColOUT As Long, myXlonEndColOUT As Long, _
            myXlonRowsOUT As Long, myXlonColsOUT As Long, _
            myXlonDscrtDataCntOUT As Long, _
            myZobjDscrtDataCellOUT() As Object, myZvarShtDscrtDataOUT() As Variant, _
            ByVal myXlonSrsDataOptnIN As Long, _
            ByVal myXlonRngOptnIN As Long, _
            ByVal myXbisByVrntIN As Boolean, ByVal myXbisGetCmntIN As Boolean)
    
'//入力変数を初期化
    myXlonSrsDataOptn = Empty
    
    myXlonRngOptn = Empty
    myXbisByVrnt = False: myXbisGetCmnt = False
    
'//入力変数を取り込み
    myXlonSrsDataOptn = myXlonSrsDataOptnIN
    
    myXlonRngOptn = myXlonRngOptnIN
    myXbisByVrnt = myXbisByVrntIN
    myXbisGetCmnt = myXbisGetCmntIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
    Set myXobjBookOUT = Nothing: myXstrShtNameOUT = Empty: myXlonShtNoOUT = Empty

    myXlonSrsDataRowCntOUT = Empty: myXlonSrsDataColCntOUT = Empty
    Erase myZstrShtSrsDataOUT: Erase myZvarShtSrsDataOUT: Erase myZstrCmntDataOUT
    myXlonBgnRowOUT = Empty: myXlonEndRowOUT = Empty
    myXlonBgnColOUT = Empty: myXlonEndColOUT = Empty
    myXlonRowsOUT = Empty: myXlonColsOUT = Empty

    myXlonDscrtDataCntOUT = Empty
    Erase myZobjDscrtDataCellOUT: Erase myZvarShtDscrtDataOUT
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    Set myXobjBookOUT = myXobjBook
    myXstrShtNameOUT = myXstrShtName
    myXlonShtNoOUT = myXlonShtNo
    
    If myXlonSrsDataOptn = 1 Then
        myXlonSrsDataRowCntOUT = myXlonSrsDataRowCnt
        myXlonSrsDataColCntOUT = myXlonSrsDataColCnt
        myZstrShtSrsDataOUT() = myZstrShtSrsData()
        myZvarShtSrsDataOUT() = myZvarShtSrsData()
        myZstrCmntDataOUT() = myZstrCmntData()
        myXlonBgnRowOUT = myXlonBgnRow
        myXlonEndRowOUT = myXlonEndRow
        myXlonBgnColOUT = myXlonBgnCol
        myXlonEndColOUT = myXlonEndCol
        myXlonRowsOUT = myXlonRows
        myXlonColsOUT = myXlonCols
        
    ElseIf myXlonSrsDataOptn = 2 Then
        myXlonDscrtDataCntOUT = myXlonDscrtDataCnt
        myZobjDscrtDataCellOUT() = myZobjDscrtDataCell()
        myZvarShtDscrtDataOUT() = myZvarShtDscrtData()
        
    Else
    End If
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:制御用変数を設定
    Call setControlVariables
    
'//S:エクセルシート上のデータ範囲を選択してデータを取得
    Select Case myXlonSrsDataOptn
    '//連続範囲
        Case 1
            Call setControlVariables1
            Call instCSlctShtSrsData
        
    '//不連続範囲
        Case 2
            Call setControlVariables2
            Call instCSlctShtDscrtData
        
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    Set myXobjBook = Nothing: myXstrShtName = Empty: myXlonShtNo = Empty
    
    myXlonSrsDataRowCnt = Empty: myXlonSrsDataColCnt = Empty
    Erase myZstrShtSrsData: Erase myZvarShtSrsData: Erase myZstrCmntData
    myXlonBgnRow = Empty: myXlonEndRow = Empty
    myXlonBgnCol = Empty: myXlonEndCol = Empty
    myXlonRows = Empty: myXlonCols = Empty
    
    myXlonDscrtDataCnt = Empty: Erase myZobjDscrtDataCell: Erase myZvarShtDscrtData
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
    
'    If myXlonSrsDataOptn < 1 And myXlonSrsDataOptn > 2 Then GoTo ExitPath
    
'    If myXlonRngOptn < 0 And myXlonRngOptn > 3 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
    
    myXlonSrsDataOptn = 1
    'myXlonSrsDataOptn = 1 : 連続データを取得する
    'myXlonSrsDataOptn = 2 : 不連続データを取得する
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables1()
    
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : 選択範囲
    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
    'myXlonRngOptn = 3  : 全データ範囲
    
    myXbisByVrnt = False
    'myXbisByVrnt = False : シートデータをStringで取得する
    'myXbisByVrnt = True  : シートデータをVariantで取得する
    
    myXbisGetCmnt = True
    'myXbisGetCmnt = False : コメントを取得しない
    'myXbisGetCmnt = True  : コメントを取得する
    
    myXstrInptBxPrmpt = ""
    myXstrInptBxTtl = ""
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    myXbisByVrnt = False
    'myXbisByVrnt = False : シートデータをStringで取得する
    'myXbisByVrnt = True  : シートデータをVariantで取得する
    
    myXbisGetCmnt = True
    'myXbisGetCmnt = False : コメントを取得しない
    'myXbisGetCmnt = True  : コメントを取得する
    
    myXstrInptBxPrmpt = ""
    myXstrInptBxTtl = ""
    
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

'◆ClassProc名_シート上の連続範囲を指定してその範囲のデータと情報を取得する
Private Sub instCSlctShtSrsData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSlctShtSrsData As CSlctShtSrsData: Set myXinsSlctShtSrsData = New CSlctShtSrsData
    With myXinsSlctShtSrsData
    '//クラス内変数への入力
        .letRngOptn = myXlonRngOptn
        .letByVrnt = myXbisByVrnt
        .letGetCmnt = myXbisGetCmnt
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonSrsDataRowCnt = .getDataRowCnt
        myXlonSrsDataColCnt = .getDataColCnt
        If myXlonSrsDataRowCnt <= 0 Or myXlonSrsDataColCnt <= 0 Then GoTo ExitPath
        i = myXlonSrsDataRowCnt + Lo - 1: j = myXlonSrsDataColCnt + Lo - 1
        ReDim myZstrShtSrsData(i, j) As String
        ReDim myZvarShtSrsData(i, j) As Variant
        ReDim myZstrCmntData(i, j) As String
        Lc = .getOptnBase
        If myXbisByVrnt = False Then
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZstrShtSrsData(i + Lo - 1, j + Lo - 1) _
                        = .getStrShtDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        Else
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZvarShtSrsData(i + Lo - 1, j + Lo - 1) _
                        = .getVarShtDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        End If
        If myXbisGetCmnt = True Then
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZstrCmntData(i + Lo - 1, j + Lo - 1) _
                        = .getCmntDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        End If
        Set myXobjBook = .getBook
        myXstrShtName = .getShtName
        myXlonShtNo = .getShtNo
        myXlonBgnRow = .getBgnEndRowCol(1, 1)
        myXlonEndRow = .getBgnEndRowCol(2, 1)
        myXlonBgnCol = .getBgnEndRowCol(1, 2)
        myXlonEndCol = .getBgnEndRowCol(2, 2)
        myXlonRows = .getBgnEndRowCol(1, 0)
        myXlonCols = .getBgnEndRowCol(0, 1)
    End With
    Set myXinsSlctShtSrsData = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSlctShtSrsData = Nothing
End Sub

'◆ClassProc名_シート上の不連続範囲を指定してその範囲のデータと情報を取得する
Private Sub instCSlctShtDscrtData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsSlctShtDscrtData As CSlctShtDscrtData
    Set myXinsSlctShtDscrtData = New CSlctShtDscrtData
    With myXinsSlctShtDscrtData
    '//クラス内変数への入力
        .letByVrnt = myXbisByVrnt
        .letGetCmnt = myXbisGetCmnt
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonDscrtDataCnt = .getDataCnt
        If myXlonDscrtDataCnt <= 0 Then GoTo ExitPath
        i = myXlonDscrtDataCnt + Lo - 1
        ReDim myZobjDscrtDataCell(i) As Object
        ReDim myZvarShtDscrtData(i, Lo + 3) As Variant
        Lc = .getOptnBase
        For i = 1 To myXlonDscrtDataCnt
            Set myZobjDscrtDataCell(i + Lo - 1) = .getDataCellAry(i + Lc - 1)
        Next i
        For i = 1 To myXlonDscrtDataCnt
            myZvarShtDscrtData(i + Lo - 1, Lo + 0) = .getShtDataAry(i + Lc - 1, Lc + 0)
            myZvarShtDscrtData(i + Lo - 1, Lo + 1) = .getShtDataAry(i + Lc - 1, Lc + 1)
            myZvarShtDscrtData(i + Lo - 1, Lo + 2) = .getShtDataAry(i + Lc - 1, Lc + 2)
        Next i
        If myXbisGetCmnt = True Then
            For i = 1 To myXlonDscrtDataCnt
                myZvarShtDscrtData(i + Lo - 1, Lo + 3) = .getShtDataAry(i + Lc - 1, Lo + 3)
            Next i
        End If
        Set myXobjBook = .getBook
        myXstrShtName = .getShtName
        myXlonShtNo = .getShtNo
    End With
    Set myXinsSlctShtDscrtData = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSlctShtDscrtData = Nothing
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
'Private Sub setControlVariables()
'    myXlonSrsDataOptn = 1
'    'myXlonSrsDataOptn = 1 : 連続データを取得する
'    'myXlonSrsDataOptn = 2 : 不連続データを取得する
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables1()
'    myXlonRngOptn = 0
'    'myXlonRngOptn = 0  : 選択範囲
'    'myXlonRngOptn = 1  : 選択位置から最終行までの範囲
'    'myXlonRngOptn = 2  : 選択位置から最終列までの範囲
'    'myXlonRngOptn = 3  : 全データ範囲
'    myXbisByVrnt = False
'    'myXbisByVrnt = False : シートデータをStringで取得する
'    'myXbisByVrnt = True  : シートデータをVariantで取得する
'    myXbisGetCmnt = True
'    'myXbisGetCmnt = False : コメントを取得しない
'    'myXbisGetCmnt = True  : コメントを取得する
'    myXstrInptBxPrmpt = ""
'    myXstrInptBxTtl = ""
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    myXbisByVrnt = False
'    'myXbisByVrnt = False : シートデータをStringで取得する
'    'myXbisByVrnt = True  : シートデータをVariantで取得する
'    myXbisGetCmnt = True
'    'myXbisGetCmnt = False : コメントを取得しない
'    'myXbisGetCmnt = True  : コメントを取得する
'    myXstrInptBxPrmpt = ""
'    myXstrInptBxTtl = ""
'End Sub
'◆ModuleProc名_シート上の範囲を指定してその範囲のデータと情報を取得する
Private Sub callxRefSlctShtData()
'  Dim myXlonSrsDataOptn As Long, myXlonRngOptn As Long, _
'        myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
'  Dim myXlonSrsDataRowCnt As Long, myXlonSrsDataColCnt As Long, _
'        myZstrShtSrsData() As String, myZvarShtSrsData() As Variant, _
'        myZstrCmntData() As String, _
'        myXlonBgnRow As Long, myXlonEndRow As Long, _
'        myXlonBgnCol As Long, myXlonEndCol As Long, _
'        myXlonRows As Long, myXlonCols As Long
'    'myZstrShtSrsData(i, j) : 取得文字列
'    'myZvarShtSrsData(i, j) : 取得文字列
'    'myZstrCmntData(i, j) : 取得コメント
'  Dim myXlonDscrtDataCnt As Long, _
'        myZobjDscrtDataCell() As Object, myZvarShtDscrtData() As Variant
'    'myZvarShtDscrtData(i, 1) = Row
'    'myZvarShtDscrtData(i, 2) = Column
'    'myZvarShtDscrtData(i, 3) = SheetData
'    'myZvarShtDscrtData(i, 4) = CommentData
    Call xRefSlctShtData.callProc( _
            myXbisCmpltFlag, _
            myXobjBook, myXstrShtName, myXlonShtNo, _
            myXlonSrsDataRowCnt, myXlonSrsDataColCnt, _
            myZstrShtSrsData, myZvarShtSrsData, myZstrCmntData, _
            myXlonBgnRow, myXlonEndRow, _
            myXlonBgnCol, myXlonEndCol, _
            myXlonRows, myXlonCols, _
            myXlonDscrtDataCnt, myZobjDscrtDataCell, myZvarShtDscrtData, _
            myXlonSrsDataOptn, myXlonRngOptn, myXbisByVrnt, myXbisGetCmnt)
    Call variablesOfxRefSlctShtData(myXlonSrsDataRowCnt, myZstrShtSrsData)   'Debug.Print
'    Call variablesOfxRefSlctShtData(myXlonDscrtDataCnt, myZvarShtDscrtData)  'Debug.Print
End Sub
Private Sub variablesOfxRefSlctShtData( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefSlctShtData内から出力した変数の内容確認
    Debug.Print "データ数: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim i As Long, j As Long
    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
            Debug.Print "データ" & i & "," & j & ": " & myZvarField(i, j)
        Next j
    Next i
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefSlctShtData()
'//xRefSlctShtDataモジュールのモジュールメモリのリセット処理
    Call xRefSlctShtData.resetConstant
End Sub
