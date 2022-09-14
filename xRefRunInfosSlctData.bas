Attribute VB_Name = "xRefRunInfosSlctData"
'Includes xRefSlctShtData
'Includes xRefRunInfos
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̏���I�����ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunInfosSlctData"
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
  Private myXlonSrsDataOptn As Long, myXlonRngOptn As Long, _
            myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean
        
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : �����
    

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonSrsDataOptn = myXlonRngOptn = Empty
    myXbisByVrnt = False: myXbisGetCmnt = False
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
    '����:  '��ModuleProc��_�V�[�g��͈̔͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
            '��ModuleProc��_�������ɑ΂��ĘA�����������{
    '�o��: -
    
'//�������s
    Call callxRefRunInfosSlctData
    
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
    Call setControlVariables
    Call setControlVariables1
    Call setControlVariables2
    
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
Private Sub setControlVariables()
    myXlonSrsDataOptn = 1
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s�A���f�[�^���擾����
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
    myXbisByVrnt = False
    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
    myXbisGetCmnt = False
    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
    'myXbisGetCmnt = True  : �R�����g���擾����
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    myXbisByVrnt = False
    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
    myXbisGetCmnt = False
    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
    'myXbisGetCmnt = True  : �R�����g���擾����
End Sub

'SnsP_�V�[�g��̏����擾����
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
    'myZstrShtSrsData(i, j) : �擾������
    'myZvarShtSrsData(i, j) : �擾������
    'myZstrCmntData(i, j) : �擾�R�����g
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

'��ModuleProc��_�V�[�g��̏���I�����ĘA�����������{����
Private Sub callxRefRunInfosSlctData()
  Dim myXbisCompFlag As Boolean
    Call xRefRunInfosSlctData.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunInfosSlctData()
'//xRefRunInfosSlctData���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunInfosSlctData.resetConstant
End Sub
