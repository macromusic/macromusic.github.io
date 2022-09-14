Attribute VB_Name = "xRefShtSrsDataLst"
'Includes CSeriesData
'Includes CSeriesAry
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上の連続するデータ範囲を取得する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefShtSrsDataLst"
  Private Const meMlonExeNum As Long = 0
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : 取得文字列
  
  Private myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
    'myZstrSrsAry(i, j) : 取得文字列
  
'//入力制御信号
  Private myXlonDataListOptn As Long
    'myXlonSrsDataOptn = 1 : 連続データを取得する
    'myXlonSrsDataOptn = 2 : 行列データを取得する
  
'//入力データ
  Private myXbisRowDrctn As Boolean
  Private myXlonBgnRow As Long, myXlonBgnCol As Long, myXobjSrchSheet As Object
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
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
    Call callxRefShtSrsDataLst
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonSrsDataCntOUT As Long, myZstrSrsDataOUT() As String, _
            myXlonSrsRowCntOUT As Long, myXlonSrsColCntOUT As Long, _
            myZstrSrsAryOUT() As String, _
            ByVal myXlonDataListOptnIN As Long, _
            ByVal myXbisRowDrctnIN As Boolean, _
            ByVal myXlonBgnRowIN As Long, ByVal myXlonBgnColIN As Long, _
            ByVal myXobjSrchSheetIN As Object)
    
'//入力変数を初期化
    myXlonDataListOptn = Empty
    
    myXbisRowDrctn = False
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    Set myXobjSrchSheet = Nothing
    
'//入力変数を取り込み
    myXlonDataListOptn = myXlonDataListOptnIN
    
    myXbisRowDrctn = myXbisRowDrctnIN
    myXlonBgnRow = myXlonBgnRowIN
    myXlonBgnCol = myXlonBgnColIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
    myXlonSrsDataCntOUT = Empty: Erase myZstrSrsDataOUT
    myXlonSrsRowCntOUT = Empty: myXlonSrsColCntOUT = Empty: Erase myZstrSrsAryOUT
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonDataListOptn = 1 Then
        myXlonSrsDataCntOUT = myXlonSrsDataCnt
        myZstrSrsDataOUT() = myZstrSrsData()
        
    ElseIf myXlonDataListOptn = 2 Then
        myXlonSrsRowCntOUT = myXlonSrsRowCnt
        myXlonSrsColCntOUT = myXlonSrsColCnt
        myZstrSrsAryOUT() = myZstrSrsAry()
        
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
    
'//S:シート上の文字列を検索してデータを取得
    Select Case myXlonDataListOptn
    '//連続データを取得
        Case 1
            Call setControlVariables1
            Call instCSeriesData
        
    '//行列データを取得
        Case 2
            Call setControlVariables2
            Call instCSeriesAry
        
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
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
    myXlonSrsRowCnt = Empty: myXlonSrsColCnt = Empty: Erase myZstrSrsAry
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
    
'    If myXlonDataListOptn < 1 And myXlonDataListOptn > 2 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables()
    
    myXlonDataListOptn = 1
    'myXlonSrsDataOptn = 1 : 連続データを取得する
    'myXlonSrsDataOptn = 2 : 行列データを取得する
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables1()
    
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : 行方向のみを検索
    'myXbisRowDrctn = False : 列方向のみを検索
    
    myXlonBgnRow = 8
    myXlonBgnCol = 2
    
  Dim myXlonSrchShtNo As Long
    myXlonSrchShtNo = 2
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    
    myXlonBgnRow = 8
    myXlonBgnCol = 2
    
  Dim myXlonSrchShtNo As Long
    myXlonSrchShtNo = 2
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    
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

'◆ClassProc名_シート上の連続するデータ範囲を取得する
Private Sub instCSeriesData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
    With myXinsSeriesData
    '//クラス内変数への入力
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
        .letRowDrctn = True
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonSrsDataCnt = .getSrsDataCnt
        If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
        k = myXlonSrsDataCnt + Lo - 1
        ReDim myZstrSrsData(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonSrsDataCnt
            myZstrSrsData(k + Lo - 1) = .getSrsDataAry(k + Lc - 1)
        Next k
    End With
    Set myXinsSeriesData = Nothing
    Set myXobjSrchSheet = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSeriesData = Nothing
    Set myXobjSrchSheet = Nothing
End Sub

'◆ClassProc名_シート上の連続するデータ範囲を行列で取得する
Private Sub instCSeriesAry()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSeriesData As CSeriesAry: Set myXinsSeriesData = New CSeriesAry
    With myXinsSeriesData
    '//クラス内変数への入力
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
    '//クラス内プロシージャの実行とクラス内変数からの出力
        .exeProc
        myXlonSrsRowCnt = .getSrsRowCnt
        myXlonSrsColCnt = .getSrsColCnt
        If myXlonSrsRowCnt <= 0 Or myXlonSrsColCnt <= 0 Then GoTo ExitPath
        i = myXlonSrsRowCnt + Lo - 1: j = myXlonSrsColCnt + Lo - 1
        ReDim myZstrSrsAry(i, j) As String
        Lc = .getOptnBase
        For j = 1 To myXlonSrsColCnt
            For i = 1 To myXlonSrsRowCnt
                myZstrSrsAry(i + Lo - 1, j + Lo - 1) _
                    = .getSrsDataAry(i + Lc - 1, j + Lc - 1)
            Next i
        Next j
    End With
    Set myXinsSeriesData = Nothing
    Set myXobjSrchSheet = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSeriesData = Nothing
    Set myXobjSrchSheet = Nothing
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
'    myXlonDataListOptn = 1
'    'myXlonSrsDataOptn = 1 : 連続データを取得する
'    'myXlonSrsDataOptn = 2 : 行列データを取得する
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables1()
'    myXbisRowDrctn = True
'    'myXbisRowDrctn = True  : 行方向のみを検索
'    'myXbisRowDrctn = False : 列方向のみを検索
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'End Sub
''SetP_制御用変数を設定する
'Private Sub setControlVariables2()
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'End Sub
'◆ModuleProc名_シート上の連続するデータ範囲を取得する
Private Sub callxRefShtSrsDataLst()
'  Dim myXlonDataListOptn As Long, myXbisRowDrctn As Boolean, _
'        myXlonBgnRow As Long, myXlonBgnCol As Long, myXobjSrchSheet As Object
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonSrsDataCnt As Long, myZstrSrsData() As String
'    'myZstrSrsData(k) : 取得文字列
'  Dim myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
'    'myZstrSrsAry(i, j) : 取得文字列
    Call xRefShtSrsDataLst.callProc( _
            myXbisCmpltFlag, _
            myXlonSrsDataCnt, myZstrSrsData, _
            myXlonSrsRowCnt, myXlonSrsColCnt, myZstrSrsAry, _
            myXlonDataListOptn, _
            myXbisRowDrctn, myXlonBgnRow, myXlonBgnCol, myXobjSrchSheet)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefShtSrsDataLst()
'//xRefShtSrsDataLstモジュールのモジュールメモリのリセット処理
    Call xRefShtSrsDataLst.resetConstant
End Sub
