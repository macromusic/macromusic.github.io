Attribute VB_Name = "xRefRunFilesShtFile"
'Includes xRefGetShtFileLst
'Includes xRefRunFiles
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̃t�@�C���p�X��I�����ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunFilesShtFile"
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
  Private myXbisByDscrt As Boolean
  Private myXlonRngOptn As Long
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
    'myXbisByDscrt = False : �V�[�g��̘A���͈͂��w�肵�Ď擾����
    'myXbisByDscrt = True  : �V�[�g��̕s�A���͈͂��w�肵�Ď擾����
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
  
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
    'myZstrOrgFilePath(i) : ���t�@�C���p�X

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisByDscrt = False
    myXlonRngOptn = Empty
    myXstrInptBxPrmpt = Empty: myXstrInptBxTtl = Empty
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
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
    '����:  '��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�f�[�^��I�����ăp�X�ꗗ���擾����
            '��ModuleProc��_�����t�@�C���ɑ΂��ĘA�����������{����
    '�o��: -
    
'//�������s
    Call callxRefRunFilesShtFile
    
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
    
'//S:�����t�@�C���p�X���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:�����t�@�C���ɑ΂��ĘA�����������{
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
    myXbisByDscrt = False
    'myXbisByDscrt = False : �V�[�g��̘A���͈͂��w�肵�Ď擾����
    'myXbisByDscrt = True  : �V�[�g��̕s�A���͈͂��w�肵�Ď擾����
    myXlonRngOptn = 0
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
    myXstrInptBxPrmpt = "�����������t�@�C���p�X��I�����ĉ������B"
    myXstrInptBxTtl = "�t�@�C���p�X�̑I��"
End Sub

'SnsP_�����t�@�C���p�X���擾����
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
    Call xRefGetShtFileLst.callProc( _
            myXlonFileCnt, myZstrFileName, myZstrFilePath)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgFileCnt = myXlonFileCnt
    myZstrOrgFilePath() = myZstrFilePath()
    
    Erase myZstrFileName: Erase myZstrFilePath
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�����t�@�C���ɑ΂��ĘA�����������{����
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeFileCnt As Long, _
        myZstrExeFileName() As String, myZstrExeFilePath() As String
    'myZstrExeFileName(i) : ���s�t�@�C����
    'myZstrExeFilePath(i) : ���s�t�@�C���p�X
    
    Call xRefRunFiles.callProc( _
            myXbisCompFlag, _
            myXlonExeFileCnt, myZstrExeFileName, myZstrExeFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath)
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

'��ModuleProc��_�V�[�g��̃t�@�C���p�X��I�����ĘA�����������{����
Private Sub callxRefRunFilesShtFile()
  Dim myXbisCompFlag As Boolean
    Call xRefRunFilesShtFile.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesShtFile()
'//xRefRunFilesShtFile���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunFilesShtFile.resetConstant
End Sub
