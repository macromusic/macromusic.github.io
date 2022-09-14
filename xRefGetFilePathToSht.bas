Attribute VB_Name = "xRefGetFilePathToSht"
'Includes CSlctFilePath
'Includes CVrblToSht
'Includes PfncbisCheckFolderExist
'Includes PfncstrCheckAndGetFilesParentFolder
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�t�@�C����I�����Ă��̃p�X���ʒu���w�肵�ăV�[�g�ɏ����o��
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefGetFilePathToSht"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, _
            myZstrFileName() As String, myZstrFilePath() As String
    'myZstrFileName(i) : �t�@�C����
    'myZstrFilePath(i) : �t�@�C���p�X
  Private myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
  
'//���͐���M��
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�H���_�p�X�^�����G�N�Z���V�[�g�ɏ����o��
    
'//���̓f�[�^
  Private myXstrDfltFldrPath As String, myXstrExtsn As String
  Private myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long
    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
    'myXbisFltrClr = False : �t�@�C���t�B���^�������������ɒǉ�����
    'myXbisFltrClr = True  : �t�@�C���t�B���^������������
    'myXlonFltrIndx = 1�` : �t�@�C���t�B���^�̏����I��
  Private myXlonIniView As Long, myXbisMultSlct As Boolean
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
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  Private myXbisCurDirON As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXstrDfltDirPath As String
  Private myZvarPstData As Variant, myXobjPstdCell As Object
  
'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    myXbisCurDirON = False
    
    myXstrDfltDirPath = Empty
    myZvarPstData = Empty: Set myXobjPstdCell = Nothing
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
    Call callxRefGetFilePathToSht
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, myXobjDirPstdCellOUT As Object, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrDfltFldrPathIN As String, ByVal myXstrExtsnIN As String, _
            ByRef myZstrAddFltrIN() As String, _
            ByVal myXbisFltrClrIN As Boolean, ByVal myXlonFltrIndxIN As Long, _
            ByVal myXlonIniViewIN As Long, ByVal myXbisMultSlctIN As Boolean)
    
'//���͕ϐ���������
    myXlonOutputOptn = Empty
    
    myXstrDfltFldrPath = Empty: myXstrExtsn = Empty
    Erase myZstrAddFltr: myXbisFltrClr = False: myXlonFltrIndx = Empty
    myXlonIniView = Empty: myXbisMultSlct = False

'//���͕ϐ�����荞��
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrDfltFldrPath = myXstrDfltFldrPathIN
    myXstrExtsn = myXstrExtsnIN
    myZstrAddFltr() = myZstrAddFltrIN()
    myXbisFltrClr = myXbisFltrClrIN
    myXlonFltrIndx = myXlonFltrIndxIN
    myXlonIniView = myXlonIniViewIN
    myXbisMultSlct = myXbisMultSlctIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    myXlonFileCntOUT = Empty: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdCellOUT = Nothing: Set myXobjDirPstdCellOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    myXlonFileCntOUT = myXlonFileCnt
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    Set myXobjFilePstdCellOUT = myXobjFilePstdCell
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell

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
    
'//S:�t�@�C����I�����Ă��̃p�X���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:�t�@�C���p�X���V�[�g�ɏ����o��
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonFileCnt = Empty: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
End Sub

'RemP_���W���[���������ɕۑ������ϐ������o��
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
    If meMlonExeNum > 0 Then myXbisCurDirON = True
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_���͕ϐ����e���m�F����
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'    If myXlonOutputOptn < 0 And myXlonOutputOptn > 3 Then myXlonOutputOptn = 1
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
    
'    myXbisCurDirON = False
'    'myXbisCurDirON = False : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肵�Ȃ�
'    'myXbisCurDirON = True  : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肷��

End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    
    myXlonOutputOptn = 1
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�H���_�p�X�^�����G�N�Z���V�[�g�ɏ����o��

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
    
    myXbisMultSlct = True
    'myXbisMultSlct = False : �����̃t�@�C����I��s�\
    'myXbisMultSlct = True  : �����̃t�@�C����I���\

End Sub

'SnsP_�t�@�C����I�����Ă��̃p�X���擾
Private Sub snsProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt As String _
        = "�_�C�A���O�{�b�N�X���\������܂��̂ŁA�t�@�C����I�����ĉ������B"
    
    If myXbisCurDirON = True Then myXstrDfltDirPath = CurDir
    If PfncbisCheckFolderExist(myXstrDfltDirPath) = False Then _
        myXstrDfltDirPath = myXstrDfltFldrPath

'//�t�@�C����I�����Ă��̃p�X���擾
    MsgBox coXstrMsgBxPrmpt
    Call instCSlctFilePath
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_
Private Sub prsProc()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�t�@�C���p�X���V�[�g�ɏ����o��
Private Sub runProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt1 As String _
        = "�t�@�C���p�X��\��t����ʒu���w�肵�ĉ������B"
  Const coXstrMsgBxPrmpt2 As String _
        = "�t�@�C������\��t����ʒu���w�肵�ĉ������B"
  Const coXstrMsgBxPrmpt3 As String _
        = "�f�B���N�g���p�X��\��t����ʒu���w�肵�ĉ������B"
   
    If myXlonOutputOptn = 0 Then Exit Sub
    
'//�t�@�C���p�X�ꗗ�̐e�t�H���_�����ꂩ�m�F���ē���ł���ΐe�t�H���_�p�X���擾
  Dim myXstrPrntPath As String
    myXstrPrntPath = PfncstrCheckAndGetFilesParentFolder(myZstrFilePath)
    If myXstrPrntPath = "" Then myXlonOutputOptn = 1
        
'//�t�@�C���p�X���V�[�g�ɏ����o�����@�ŕ���
  Dim myXbisPstFlag As Boolean
    If myXlonOutputOptn = 2 Then
    '//�t�@�C�����������o���ꍇ
        myZvarPstData = myZstrFileName
        MsgBox coXstrMsgBxPrmpt2
        
    ElseIf myXlonOutputOptn = 3 Then
    '//�e�t�H���_�ɉ����ď����o���ꍇ
    
    '//�f�B���N�g���p�X���G�N�Z���V�[�g�ɏ����o��
        myZvarPstData = myXstrPrntPath
        MsgBox coXstrMsgBxPrmpt3
        
        Call instCVrblToSht(myXbisPstFlag)
        If myXbisPstFlag = False Then GoTo ExitPath
        Set myXobjDirPstdCell = myXobjPstdCell
        
    '//�t�@�C�������G�N�Z���V�[�g�ɏ����o��
        myZvarPstData = myZstrFileName
        MsgBox coXstrMsgBxPrmpt2
        
    Else
    '//�t�@�C���p�X�������o���ꍇ
        myZvarPstData = myZstrFilePath
        MsgBox coXstrMsgBxPrmpt1
        
    End If
    
'//�t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    Call instCVrblToSht(myXbisPstFlag)
    If myXbisPstFlag = False Then GoTo ExitPath
    Set myXobjFilePstdCell = myXobjPstdCell
    
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

'��ClassProc��_�t�@�C����I�����Ă��̃p�X���擾����
Private Sub instCSlctFilePath()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsFilePath As CSlctFilePath: Set myXinsFilePath = New CSlctFilePath
    With myXinsFilePath
    '//�N���X���ϐ��ւ̓���
        .letFDType = msoFileDialogFilePicker
        .letDfltFldrPath = myXstrDfltDirPath
        .letDfltFilePath = ""
        .letExtsn = myXstrExtsn
        .letAddFltr = myZstrAddFltr
        .letFltrClr = myXbisFltrClr
        .letFltrIndx = myXlonFltrIndx
        .letIniView = myXlonIniView
        .letMultSlct = myXbisMultSlct
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonFileCnt = .getFileCnt
        If myXlonFileCnt <= 0 Then GoTo JumpPath
        k = myXlonFileCnt + Lo - 1
        ReDim myZstrFileName(k) As String
        ReDim myZstrFilePath(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFileCnt
            myZstrFileName(k + Lo - 1) = .getFileNameAry(k + Lc - 1)
            myZstrFilePath(k + Lo - 1) = .getFilePathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsFilePath = Nothing
End Sub

'��ClassProc��_�ϐ������G�N�Z���V�[�g�ɏ����o��
Private Sub instCVrblToSht(myXbisCompFlag As Boolean)
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//�N���X���ϐ��ւ̓���
        .letVrbl = myZvarPstData
        Set .setPstFrstCell = Nothing
        .letInptBxOFF = False
        .letEachWrtON = False
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        myXbisCompFlag = .fncbisCmpltFlag
        Set myXobjPstdCell = .getPstdRng
    End With
    Set myXinsVrblToSht = Nothing
End Sub

'===============================================================================================

 '��^�e_�w��t�H���_�̑��݂��m�F����
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '��^�e_�t�@�C���p�X�ꗗ�̐e�t�H���_�����ꂩ�m�F���ē���ł���ΐe�t�H���_�p�X���擾����
Private Function PfncstrCheckAndGetFilesParentFolder( _
            ByRef myZstrOrgFilePath() As String) As String
    PfncstrCheckAndGetFilesParentFolder = Empty
'//�t�@�C���̐e�t�H���_���擾
  Dim myXstrTmpFile As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgFilePath): myXstrTmpFile = myZstrOrgFilePath(L)
    On Error GoTo 0
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrPrntPath As String
    myXstrPrntPath = myXobjFSO.GetParentFolderName(myXstrTmpFile)
'//�S�t�@�C���̐e�t�H���_�����ꂩ�m�F
  Dim myXvarTmp As Variant, myXstrTmp As String
    For Each myXvarTmp In myZstrOrgFilePath
        myXstrTmp = myXobjFSO.GetParentFolderName(myXvarTmp)
        If myXstrPrntPath <> myXstrTmp Then GoTo ExitPath
    Next myXvarTmp
    PfncstrCheckAndGetFilesParentFolder = myXstrPrntPath
ExitPath:
    Set myXobjFSO = Nothing
End Function

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
'Private Sub setControlVariables1()
'    myXlonOutputOptn = 1
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXstrDfltFldrPath = "1"
'    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
'    myXstrExtsn = "pdf"
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'  Dim myXlonAddFltrCnt As Long
'    myXlonAddFltrCnt = 1
'    ReDim myZstrAddFltr(myXlonAddFltrCnt + L - 1, L + 1) As String
'    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
'    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
'  Dim i As Long: i = L - 1
'    i = i + 1   'i = 1
'    myZstrAddFltr(i, L + 0) = "PDF�t�@�C��"
'    myZstrAddFltr(i, L + 1) = "*.pdf"
'    myXbisFltrClr = False
'    'myXbisFltrClr = False : �t�@�C���t�B���^�������������ɒǉ�����
'    'myXbisFltrClr = True  : �t�@�C���t�B���^������������
'    myXlonFltrIndx = 1
'    'myXlonFltrIndx = 1�` : �t�@�C���t�B���^�̏����I��
'    myXlonIniView = msoFileDialogViewList
'    'myXlonIniView = msoFileDialogViewDetails    : �t�@�C�����ڍ׏��Ƌ��Ɉꗗ�\��
'    'myXlonIniView = msoFileDialogViewLargeIcons : �t�@�C����傫���A�C�R���ŕ\��
'    'myXlonIniView = msoFileDialogViewList       : �t�@�C�����ڍ׏��Ȃ��ňꗗ�\��
'    'myXlonIniView = msoFileDialogViewPreview    : �t�@�C���̈ꗗ��\�����A�I�������t�@�C�����v���r���[ �E�B���h�E�g�ɕ\��
'    'myXlonIniView = msoFileDialogViewProperties : �t�@�C���̈ꗗ��\�����A�I�������t�@�C���̃v���p�e�B���E�B���h�E�g�ɕ\��
'    'myXlonIniView = msoFileDialogViewSmallIcons : �t�@�C�����������A�C�R���ŕ\��
'    'myXlonIniView = msoFileDialogViewThumbnail  : �t�@�C�����k���\��
'    'myXlonIniView = msoFileDialogViewTiles      : �t�@�C�����A�C�R���ŕ��ׂĕ\��
'    'myXlonIniView = msoFileDialogViewWebView    : �t�@�C���� Web �\��
'    myXbisMultSlct = False
'    'myXbisMultSlct = False : �����̃t�@�C����I��s�\
'    'myXbisMultSlct = True  : �����̃t�@�C����I���\
'End Sub
'��ModuleProc��_�t�@�C����I�����Ă��̃p�X���V�[�g�ɏ����o��
Private Sub callxRefGetFilePathToSht()
'  Dim myXlonOutputOptn As Long, _
'        myXstrDfltFldrPath As String, myXstrExtsn As String, _
'        myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long, _
'        myXlonIniView As Long, myXbisMultSlct As Boolean
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
'    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
'    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
'    'myXbisFltrClr = False : �t�@�C���t�B���^�������������ɒǉ�����
'    'myXbisFltrClr = True  : �t�@�C���t�B���^������������
'    'myXlonFltrIndx = 1�` : �t�@�C���t�B���^�̏����I��
'    'myXlonIniView = msoFileDialogViewDetails    : �t�@�C�����ڍ׏��Ƌ��Ɉꗗ�\��
'    'myXlonIniView = msoFileDialogViewLargeIcons : �t�@�C����傫���A�C�R���ŕ\��
'    'myXlonIniView = msoFileDialogViewList       : �t�@�C�����ڍ׏��Ȃ��ňꗗ�\��
'    'myXlonIniView = msoFileDialogViewPreview    : �t�@�C���̈ꗗ��\�����A�I�������t�@�C�����v���r���[ �E�B���h�E�g�ɕ\��
'    'myXlonIniView = msoFileDialogViewProperties : �t�@�C���̈ꗗ��\�����A�I�������t�@�C���̃v���p�e�B���E�B���h�E�g�ɕ\��
'    'myXlonIniView = msoFileDialogViewSmallIcons : �t�@�C�����������A�C�R���ŕ\��
'    'myXlonIniView = msoFileDialogViewThumbnail  : �t�@�C�����k���\��
'    'myXlonIniView = msoFileDialogViewTiles      : �t�@�C�����A�C�R���ŕ��ׂĕ\��
'    'myXlonIniView = msoFileDialogViewWebView    : �t�@�C���� Web �\��
'    'myXbisMultSlct = False : �����̃t�@�C����I��s�\
'    'myXbisMultSlct = True  : �����̃t�@�C����I���\
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
'    'myZstrFileName(i) : �t�@�C����
'    'myZstrFilePath(i) : �t�@�C���p�X
    Call xRefGetFilePathToSht.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXlonOutputOptn, _
            myXstrDfltFldrPath, myXstrExtsn, _
            myZstrAddFltr, myXbisFltrClr, myXlonFltrIndx, _
            myXlonIniView, myXbisMultSlct)
    Call variablesOfxRefGetFilePathToSht(myXlonFileCnt, myZstrFileName)  'Debug.Print
End Sub
Private Sub variablesOfxRefGetFilePathToSht( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefGetFilePathToSht������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt = 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myZvarField) To UBound(myZvarField)
        Debug.Print "�f�[�^" & k & ": " & myZvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefGetFilePathToSht()
'//xRefGetFilePathToSht���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefGetFilePathToSht.resetConstant
End Sub
