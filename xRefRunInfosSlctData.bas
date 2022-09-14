Attribute VB_Name = "xRefRunInfosSlctData"
'Includes xRefSlctShtData
'Includes xRefRunInfos
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上の情報を選択して連続処理を実施する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRunInfosSlctData"
  Private Const meMlonExeNum As Long = 0
  
'//モジュール内定数

'//モジュール内定数_列挙体
  
'//出力制御信号
  Private myXbisCmpltFlag As Boolean
  
'//出力データ
  
'//入力制御信号
  
'//入力データ
  
'//モジュール内変数_制御信号
  Private myXbisExitFlag As Boolean
  
'//モジュール内変数_データ
  Private myXlonSrsDataOptn As Long, myXlonRngOptn As Long, _
            myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean
        
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : 元情報
    

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonSrsDataOptn = myXlonRngOptn = Empty
    myXbisByVrnt = False: myXbisGetCmnt = False
    myXlonOrgInfoCnt = Empty: Erase myZstrOrgInfo
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

'//プログラム構成
    '入力: -
    '処理:  '◆ModuleProc名_シート上の範囲を指定してその範囲のデータと情報を取得する
            '◆ModuleProc名_複数情報に対して連続処理を実施
    '出力: -
    
'//処理実行
    Call callxRefRunInfosSlctData
    
'//処理結果表示
    Select Case myXbisCmpltFlag
        Case True: MsgBox "実行完了"
        Case Else: MsgBox "異常あり", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
    
'//出力変数を初期化
    myXbisCmpltFlagOUT = False
    
'//処理実行
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//出力変数に格納
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
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
    Call setControlVariables1
    Call setControlVariables2
    
'//S:シート上の情報を取得
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:複数情報に対して連続処理を実施
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_出力変数を初期化する
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
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
    myXbisGetCmnt = False
    'myXbisGetCmnt = False : コメントを取得しない
    'myXbisGetCmnt = True  : コメントを取得する
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariables2()
    myXbisByVrnt = False
    'myXbisByVrnt = False : シートデータをStringで取得する
    'myXbisByVrnt = True  : シートデータをVariantで取得する
    myXbisGetCmnt = False
    'myXbisGetCmnt = False : コメントを取得しない
    'myXbisGetCmnt = True  : コメントを取得する
End Sub

'SnsP_シート上の情報を取得する
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
  Dim myXlonSrsDataRowCnt As Long, myXlonSrsDataColCnt As Long, _
        myZstrShtSrsData() As String, myZvarShtSrsData() As Variant, _
        myZstrCmntData() As String, _
        myXlonBgnRow As Long, myXlonEndRow As Long, _
        myXlonBgnCol As Long, myXlonEndCol As Long, _
        myXlonRows As Long, myXlonCols As Long
    'myZstrShtSrsData(i, j) : 取得文字列
    'myZvarShtSrsData(i, j) : 取得文字列
    'myZstrCmntData(i, j) : 取得コメント
  Dim myXlonDscrtDataCnt As Long, _
        myZobjDscrtDataCell() As Object, myZvarShtDscrtData() As Variant
    'myZvarShtDscrtData(i, 1) = Row
    'myZvarShtDscrtData(i, 2) = Column
    'myZvarShtDscrtData(i, 3) = SheetData
    'myZvarShtDscrtData(i, 4) = CommentData
    
    Call xRefSlctShtData.callProc( _
            myXbisCompFlag, _
            myXobjBook, myXstrShtName, myXlonShtNo, _
            myXlonSrsDataRowCnt, myXlonSrsDataColCnt, _
            myZstrShtSrsData, myZvarShtSrsData, myZstrCmntData, _
            myXlonBgnRow, myXlonEndRow, _
            myXlonBgnCol, myXlonEndCol, _
            myXlonRows, myXlonCols, _
            myXlonDscrtDataCnt, myZobjDscrtDataCell, myZvarShtDscrtData, _
            myXlonSrsDataOptn, myXlonRngOptn, myXbisByVrnt, myXbisGetCmnt)
    If myXlonSrsDataRowCnt <= 0 Or myXlonSrsDataColCnt <= 0 Then GoTo ExitPath
    If myXlonSrsDataColCnt > 1 Then GoTo ExitPath
    
    myXlonOrgInfoCnt = myXlonSrsDataRowCnt
    myZstrOrgInfo() = myZstrShtSrsData()
    
    Set myXobjBook = Nothing
    Erase myZstrShtSrsData: Erase myZvarShtSrsData: Erase myZstrCmntData
    Erase myZobjDscrtDataCell: Erase myZvarShtDscrtData
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_複数情報に対して連続処理を実施する
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeInfoCnt As Long, myZstrExeInfo() As String
    'myZstrExeInfo(i) : 実行情報
    
    Call xRefRunInfos.callProc( _
            myXbisCompFlag, myXlonExeInfoCnt, myZstrExeInfo, _
            myXlonOrgInfoCnt, myZstrOrgInfo)
    If myXlonExeFileCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExeFileName: Erase myZstrExeFilePath
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_
Private Sub runProc()
    myXbisExitFlag = False
    
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

'◆ModuleProc名_シート上の情報を選択して連続処理を実施する
Private Sub callxRefRunInfosSlctData()
  Dim myXbisCompFlag As Boolean
    Call xRefRunInfosSlctData.callProc(myXbisCompFlag)
    Debug.Print "結果: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunInfosSlctData()
'//xRefRunInfosSlctDataモジュールのモジュールメモリのリセット処理
    Call xRefRunInfosSlctData.resetConstant
End Sub
