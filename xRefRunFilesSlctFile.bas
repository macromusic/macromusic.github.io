Attribute VB_Name = "xRefRunFilesSlctFile"
'Includes xRefGetFilePathToSht
'Includes xRefRunFiles
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�t�@�C����I�����ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunFilesSlctFile"
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
  Private myXlonOutputOptn As Long, _
            myXstrDfltFldrPath As String, myXstrExtsn As String, _
            myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long, _
            myXlonIniView As Long, myXbisMultSlct As Boolean
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
    'myXbisFltrClr = False : �t�@�C���t�B���^�������������ɒǉ�����
    'myXbisFltrClr = True  : �t�@�C���t�B���^������������
    'myXlonFltrIndx = 1�` : �t�@�C���t�B���^�̏����I��
    'myXlonIniView = msoFileDialogViewDetails    : �t�@�C�����ڍ׏��Ƌ��Ɉꗗ�\��
    'myXlonIniView = msoFileDialogViewLargeIcons : �t�@�C����傫���A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewList       : �t�@�C�����ڍ׏��Ȃ��ňꗗ�\��
    'myXlonIniView = msoFileDialogViewPreview    : �t�@�C���̈ꗗ��\�����A�I�������t�@�C�����v���r���[ �E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewProperties : �t�@�C���̈ꗗ��\�����A�I�������t�@�C���̃v���p�e�B���E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewSmallIcons : �t�@�C�����������A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewThumbnail  : �t�@�C�����k���\��
    'myXlonIniView = msoFileDialogViewTiles      : �t�@�C�����A�C�R���ŕ��ׂĕ\��
    'myXlonIniView = msoFileDialogViewWebView    : �t�@�C���� Web �\��
    'myXbisMultSlct = False : �����̃t�@�C����I��s�\
    'myXbisMultSlct = True  : �����̃t�@�C����I���\
  
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
    'myZstrOrgFilePath(i) : ���t�@�C���p�X

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOutputOptn = Empty
    myXstrDfltFldrPath = Empty: myXstrExtsn = Empty
    Erase myZstrAddFltr: myXbisFltrClr = False: yXlonFltrIndx = Empty
    myXlonIniView = Empty: myXbisMultSlct = False
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
    '����:  '��ModuleProc��_�t�@�C����I�����Ă��̃p�X���ʒu���w�肵�ăV�[�g�ɏ����o��
            '��ModuleProc��_�����t�@�C���ɑ΂��ĘA�����������{����
    '�o��: -
    
'//�������s
    Call callxRefRunFilesSlctFile
    
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
    Call setControlVariables1
    Call setControlVariables2
    
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
Private Sub setControlVariables1()
    myXlonOutputOptn = 1
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    myXstrDfltFldrPath = "1"
    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
    myXstrExtsn = "pdf"
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXlonAddFltrCnt As Long
    myXlonAddFltrCnt = 1
    ReDim myZstrAddFltr(myXlonAddFltrCnt + L - 1, L + 1) As String
    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
  Dim i As Long: i = L - 1
    i = i + 1   'i = 1
    myZstrAddFltr(i, L + 0) = "PDF�t�@�C��"
    myZstrAddFltr(i, L + 1) = "*.pdf"
    myXbisFltrClr = False
    'myXbisFltrClr = False : �t�@�C���t�B���^�������������ɒǉ�����
    'myXbisFltrClr = True  : �t�@�C���t�B���^������������
    myXlonFltrIndx = 1
    'myXlonFltrIndx = 1�` : �t�@�C���t�B���^�̏����I��
    myXlonIniView = msoFileDialogViewList
    'myXlonIniView = msoFileDialogViewDetails    : �t�@�C�����ڍ׏��Ƌ��Ɉꗗ�\��
    'myXlonIniView = msoFileDialogViewLargeIcons : �t�@�C����傫���A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewList       : �t�@�C�����ڍ׏��Ȃ��ňꗗ�\��
    'myXlonIniView = msoFileDialogViewPreview    : �t�@�C���̈ꗗ��\�����A�I�������t�@�C�����v���r���[ �E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewProperties : �t�@�C���̈ꗗ��\�����A�I�������t�@�C���̃v���p�e�B���E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewSmallIcons : �t�@�C�����������A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewThumbnail  : �t�@�C�����k���\��
    'myXlonIniView = msoFileDialogViewTiles      : �t�@�C�����A�C�R���ŕ��ׂĕ\��
    'myXlonIniView = msoFileDialogViewWebView    : �t�@�C���� Web �\��
    myXbisMultSlct = False
    'myXbisMultSlct = False : �����̃t�@�C����I��s�\
    'myXbisMultSlct = True  : �����̃t�@�C����I���\
End Sub

'SnsP_�����t�@�C���p�X���擾����
Private Sub snsProc()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
    'myZstrFileName(i) : �t�@�C����
    'myZstrFilePath(i) : �t�@�C���p�X
    
    Call xRefGetFilePathToSht.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXlonOutputOptn, _
            myXstrDfltFldrPath, myXstrExtsn, _
            myZstrAddFltr, myXbisFltrClr, myXlonFltrIndx, _
            myXlonIniView, myXbisMultSlct)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgFileCnt = myXlonFileCnt
    myZstrOrgFilePath() = myZstrFilePath()
    
    Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
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

'��ModuleProc��_�t�@�C����I�����ĘA�����������{����
Private Sub callxRefRunFilesSlctFile()
  Dim myXbisCompFlag As Boolean
    Call xRefRunFilesSlctFile.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesSlctFile()
'//xRefRunFilesSlctFile���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunFilesSlctFile.resetConstant
End Sub
