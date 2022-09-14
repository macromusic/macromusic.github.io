Attribute VB_Name = "xRefShtSrsDataLst"
'Includes CSeriesData
'Includes CSeriesAry
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefShtSrsDataLst"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : �擾������
  
  Private myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
    'myZstrSrsAry(i, j) : �擾������
  
'//���͐���M��
  Private myXlonDataListOptn As Long
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s��f�[�^���擾����
  
'//���̓f�[�^
  Private myXbisRowDrctn As Boolean
  Private myXlonBgnRow As Long, myXlonBgnCol As Long, myXobjSrchSheet As Object
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_���W���[���������̃��Z�b�g
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()
    
'//�������s
    Call callxRefShtSrsDataLst
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
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
    
'//���͕ϐ���������
    myXlonDataListOptn = Empty
    
    myXbisRowDrctn = False
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    Set myXobjSrchSheet = Nothing
    
'//���͕ϐ�����荞��
    myXlonDataListOptn = myXlonDataListOptnIN
    
    myXbisRowDrctn = myXbisRowDrctnIN
    myXlonBgnRow = myXlonBgnRowIN
    myXlonBgnCol = myXlonBgnColIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    myXlonSrsDataCntOUT = Empty: Erase myZstrSrsDataOUT
    myXlonSrsRowCntOUT = Empty: myXlonSrsColCntOUT = Empty: Erase myZstrSrsAryOUT
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
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
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//S:�V�[�g��̕�������������ăf�[�^���擾
    Select Case myXlonDataListOptn
    '//�A���f�[�^���擾
        Case 1
            Call setControlVariables1
            Call instCSeriesData
        
    '//�s��f�[�^���擾
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

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
    myXlonSrsRowCnt = Empty: myXlonSrsColCnt = Empty: Erase myZstrSrsAry
End Sub

'RemP_���W���[���������ɕۑ������ϐ������o��
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_���͕ϐ����e���m�F����
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'    If myXlonDataListOptn < 1 And myXlonDataListOptn > 2 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
    
    myXlonDataListOptn = 1
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s��f�[�^���擾����
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : �s�����݂̂�����
    'myXbisRowDrctn = False : ������݂̂�����
    
    myXlonBgnRow = 8
    myXlonBgnCol = 2
    
  Dim myXlonSrchShtNo As Long
    myXlonSrchShtNo = 2
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    myXlonBgnRow = 8
    myXlonBgnCol = 2
    
  Dim myXlonSrchShtNo As Long
    myXlonSrchShtNo = 2
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    
End Sub

'checkP_�o�͕ϐ����e���m�F����
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_�g�p�����ϐ������W���[���������ɕۑ�����
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

'��ClassProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
Private Sub instCSeriesData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
    With myXinsSeriesData
    '//�N���X���ϐ��ւ̓���
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
        .letRowDrctn = True
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
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

'��ClassProc��_�V�[�g��̘A������f�[�^�͈͂��s��Ŏ擾����
Private Sub instCSeriesAry()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSeriesData As CSeriesAry: Set myXinsSeriesData = New CSeriesAry
    With myXinsSeriesData
    '//�N���X���ϐ��ւ̓���
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
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

 '��^�o_���W���[�����萔�̒l��ύX����
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

'Dummy�o_
Private Sub MsubDummy()
End Sub

'===============================================================================================

''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables()
'    myXlonDataListOptn = 1
'    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
'    'myXlonSrsDataOptn = 2 : �s��f�[�^���擾����
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables1()
'    myXbisRowDrctn = True
'    'myXbisRowDrctn = True  : �s�����݂̂�����
'    'myXbisRowDrctn = False : ������݂̂�����
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'End Sub
'��ModuleProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
Private Sub callxRefShtSrsDataLst()
'  Dim myXlonDataListOptn As Long, myXbisRowDrctn As Boolean, _
'        myXlonBgnRow As Long, myXlonBgnCol As Long, myXobjSrchSheet As Object
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonSrsDataCnt As Long, myZstrSrsData() As String
'    'myZstrSrsData(k) : �擾������
'  Dim myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
'    'myZstrSrsAry(i, j) : �擾������
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
'//xRefShtSrsDataLst���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefShtSrsDataLst.resetConstant
End Sub
