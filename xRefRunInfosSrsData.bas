Attribute VB_Name = "xRefRunInfosSrsData"
'Includes xRefSrchShtCmnt
'Includes xRefShtSrsDataLst
'Includes xRefRunInfos
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'◆ModuleProc名_シート上の情報を取得して連続処理を実施する
'Rev.001
  
'//モジュールメモリ
  Private Const meMstrMdlName As String = "xRefRunInfosSrsData"
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
  Private myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
            myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
            myXbisInStrOptn As Boolean
  
  Private myXlonDataListOptn As Long, myXbisRowDrctn As Boolean, _
            myXlonBgnRow As Long, myXlonBgnCol As Long
        
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : 元情報

'iniP_モジュール内変数を初期化する
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonSrchShtNo = Empty: Set myXobjSrchSheet = Nothing
    myXlonShtSrchCnt = Empty: myZvarSrchCndtn = Empty
    myXbisInStrOptn = False
    
    myXlonDataListOptn = Empty: myXbisRowDrctn = False
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    
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
    '処理:  '◆ModuleProc名_シート上のデータとコメントから文字列を検索してデータと位置情報を取得する
            '◆ModuleProc名_シート上の連続するデータ範囲を取得する
            '◆ModuleProc名_複数ファイルに対して連続処理を実施する
    '出力: -
    
'//処理実行
    Call callxRefRunInfosSrsData
    
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
    Call setControlVariablesA
    Call setControlVariablesB
    Call setControlVariablesB1
    Call setControlVariablesB2
    
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
Private Sub setControlVariablesA()
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    Set myXobjSrchSheet = ActiveSheet
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 1
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : 検索文字列
    'myZvarSrchCndtn(i, 2) : オフセット行数
    'myZvarSrchCndtn(i, 3) : オフセット列数
    'myZvarSrchCndtn(i, 4) : シート上文字列検索[=0]orコメント内文字列検索[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "サブファイル一覧"
    myZvarSrchCndtn(k, L + 1) = 1
    myZvarSrchCndtn(k, L + 2) = 0
    myZvarSrchCndtn(k, L + 3) = 0
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : 指定文字列と一致する条件で検索する
    'myXbisInStrOptn = True  : 指定文字列を含む条件で検索する
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesB()
    myXlonDataListOptn = 1
    'myXlonSrsDataOptn = 1 : 連続データを取得する
    'myXlonSrsDataOptn = 2 : 行列データを取得する
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesB1()
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : 行方向のみを検索
    'myXbisRowDrctn = False : 列方向のみを検索
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
End Sub

'SetP_制御用変数を設定する
Private Sub setControlVariablesB2()
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
End Sub

'SnsP_シート上の情報を取得する
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
    'myZstrTrgtVal(i) : 取得文字列
    'myZobjTrgtRng(i) : 行列位置のセル
    
    Call xRefSrchShtCmnt.callProc( _
            myXlonTrgtValCnt, myZstrTrgtVal, myZobjTrgtRng, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn)
    If myXlonTrgtValCnt <= 0 Then Exit Sub
    
  Dim L As Long: L = LBound(myZobjTrgtRng)
    myXlonBgnRow = myZobjTrgtRng(L).Row
    myXlonBgnCol = myZobjTrgtRng(L).Column
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : 取得文字列
  Dim myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
    'myZstrSrsAry(i, j) : 取得文字列
    Call xRefShtSrsDataLst.callProc( _
            myXbisCompFlag, _
            myXlonSrsDataCnt, myZstrSrsData, _
            myXlonSrsRowCnt, myXlonSrsColCnt, myZstrSrsAry, _
            myXlonDataListOptn, _
            myXbisRowDrctn, myXlonBgnRow, myXlonBgnCol, myXobjSrchSheet)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgInfoCnt = myXlonSrsDataCnt
    myZstrOrgInfo() = myZstrSrsData()
    
    Erase myZstrTrgtVal: Erase myZobjTrgtRng
    Erase myZstrSrsData: Erase myZstrSrsAry
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

'◆ModuleProc名_複数ファイルをリストアップして連続処理を実施する
Private Sub callxRefRunInfosSrsData()
  Dim myXbisCompFlag As Boolean
    Call xRefRunInfosSrsData.callProc(myXbisCompFlag)
    Debug.Print "結果: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunInfosSrsData()
'//xRefRunInfosSrsDataモジュールのモジュールメモリのリセット処理
    Call xRefRunInfosSrsData.resetConstant
End Sub
