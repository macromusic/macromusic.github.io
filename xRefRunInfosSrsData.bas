Attribute VB_Name = "xRefRunInfosSrsData"
'Includes xRefSrchShtCmnt
'Includes xRefShtSrsDataLst
'Includes xRefRunInfos
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̏����擾���ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunInfosSrsData"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔

'//���W���[�����萔_�񋓑�
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  
'//���͐���M��
  
'//���̓f�[�^
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
            myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
            myXbisInStrOptn As Boolean
  
  Private myXlonDataListOptn As Long, myXbisRowDrctn As Boolean, _
            myXlonBgnRow As Long, myXlonBgnCol As Long
        
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : �����

'iniP_���W���[�����ϐ�������������
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

'PublicP_���W���[���������̃��Z�b�g
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()

'//�v���O�����\��
    '����: -
    '����:  '��ModuleProc��_�V�[�g��̃f�[�^�ƃR�����g���當������������ăf�[�^�ƈʒu�����擾����
            '��ModuleProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
            '��ModuleProc��_�����t�@�C���ɑ΂��ĘA�����������{����
    '�o��: -
    
'//�������s
    Call callxRefRunInfosSrsData
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariablesA
    Call setControlVariablesB
    Call setControlVariablesB1
    Call setControlVariablesB2
    
'//S:�V�[�g��̏����擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:�������ɑ΂��ĘA�����������{
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

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
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
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesA()
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    Set myXobjSrchSheet = ActiveSheet
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 1
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : ����������
    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "�T�u�t�@�C���ꗗ"
    myZvarSrchCndtn(k, L + 1) = 1
    myZvarSrchCndtn(k, L + 2) = 0
    myZvarSrchCndtn(k, L + 3) = 0
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesB()
    myXlonDataListOptn = 1
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s��f�[�^���擾����
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesB1()
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : �s�����݂̂�����
    'myXbisRowDrctn = False : ������݂̂�����
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesB2()
'    myXlonBgnRow = 8
'    myXlonBgnCol = 2
'  Dim myXlonSrchShtNo As Long
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
End Sub

'SnsP_�V�[�g��̏����擾����
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
    'myZstrTrgtVal(i) : �擾������
    'myZobjTrgtRng(i) : �s��ʒu�̃Z��
    
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
    'myZstrSrsData(k) : �擾������
  Dim myXlonSrsRowCnt As Long, myXlonSrsColCnt As Long, myZstrSrsAry() As String
    'myZstrSrsAry(i, j) : �擾������
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

'PrcsP_�������ɑ΂��ĘA�����������{����
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeInfoCnt As Long, myZstrExeInfo() As String
    'myZstrExeInfo(i) : ���s���
    
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

'��ModuleProc��_�����t�@�C�������X�g�A�b�v���ĘA�����������{����
Private Sub callxRefRunInfosSrsData()
  Dim myXbisCompFlag As Boolean
    Call xRefRunInfosSrsData.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunInfosSrsData()
'//xRefRunInfosSrsData���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunInfosSrsData.resetConstant
End Sub
