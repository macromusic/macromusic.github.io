Attribute VB_Name = "xRefSlctFldrPath"
'Includes CSlctFldrPath
'Includes CExplrAdrs
'Includes CExplrAdrsSlct
'Includes PfncbisCheckFolderExist
'Includes PfncobjGetFolder
'Includes PfixGetFolderNameInformationByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�t�H���_��I�����Ă��̃p�X���擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSlctFldrPath"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXstrFldrPath As String, myXobjFldr As Object, _
            myXstrDirPath As String, myXstrFldrName As String
  
'//���͐���M��
  Private myXlonDirSlctOptn As Long
    'myXlonDirSlctOptn = 1 : �A�N�e�B�u�u�b�N�̐e�t�H���_�p�X���擾
    'myXlonDirSlctOptn = 2 : FileDialog�I�u�W�F�N�g���g�p���ăt�H���_�p�X���擾
    'myXlonDirSlctOptn = 3 : �őO�ʂ̃G�N�X�v���[���ɕ\������Ă���t�H���_�p�X���擾
    'myXlonDirSlctOptn = 4 : �N�����̃G�N�X�v���[����I�����Ă��̃A�h���X�o�[���擾
  
'//���̓f�[�^
  Private myXstrDfltFldrPath As String
    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
  Private myXlonIniView As Long
    'myXlonIniView = msoFileDialogViewDetails    : �t�@�C�����ڍ׏��Ƌ��Ɉꗗ�\��
    'myXlonIniView = msoFileDialogViewLargeIcons : �t�@�C����傫���A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewList       : �t�@�C�����ڍ׏��Ȃ��ňꗗ�\��
    'myXlonIniView = msoFileDialogViewPreview    : �t�@�C���̈ꗗ��\�����A�I�������t�@�C�����v���r���[ �E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewProperties : �t�@�C���̈ꗗ��\�����A�I�������t�@�C���̃v���p�e�B���E�B���h�E�g�ɕ\��
    'myXlonIniView = msoFileDialogViewSmallIcons : �t�@�C�����������A�C�R���ŕ\��
    'myXlonIniView = msoFileDialogViewThumbnail  : �t�@�C�����k���\��
    'myXlonIniView = msoFileDialogViewTiles      : �t�@�C�����A�C�R���ŕ��ׂĕ\��
    'myXlonIniView = msoFileDialogViewWebView    : �t�@�C���� Web �\��
  Private myXbisExplrAdrsMsgOptn As Boolean
    'myXbisExplrAdrsMsgOptn = True  : �E�B���h�I���̊m�F���b�Z�[�W��\������
    'myXbisExplrAdrsMsgOptn = False : �E�B���h�I���̊m�F���b�Z�[�W��\�����Ȃ�
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  Private myXbisCurDirON As Boolean
    'myXbisCurDirON = False : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肵�Ȃ�
    'myXbisCurDirON = True  : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肷��

'//���W���[�����ϐ�_�f�[�^
    
'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    myXbisCurDirON = False
    
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
    Call callxRefSlctFldrPath
    
'//�������ʕ\��
    MsgBox "�擾�p�X�F" & myXstrFldrPath
    
End Sub

'PublicP_
Public Sub callProc( _
            myXstrFldrPathOUT As String, myXobjFldrOUT As Object, _
            myXstrDirPathOUT As String, myXstrFldrNameOUT As String, _
            ByVal myXlonDirSlctOptnIN As Long, _
            ByVal myXstrDfltFldrPathIN As String, ByVal myXlonIniViewIN As Long, _
            ByVal myXbisExplrAdrsMsgOptnIN As Boolean)

'//���͕ϐ���������
    myXlonDirSlctOptn = Empty
    myXstrDfltFldrPath = Empty: myXlonIniView = Empty
    myXbisExplrAdrsMsgOptn = False

'//���͕ϐ�����荞��
    myXlonDirSlctOptn = myXlonDirSlctOptnIN
    myXstrDfltFldrPath = myXstrDfltFldrPathIN
    myXlonIniView = myXlonIniViewIN
    myXbisExplrAdrsMsgOptn = myXbisExplrAdrsMsgOptnIN
    
'//�o�͕ϐ���������
    myXstrFldrPathOUT = Empty: Set myXobjFldrOUT = Nothing
    myXstrDirPathOUT = Empty: myXstrFldrNameOUT = Empty
    
'//�������s
    Call ctrProc
    If myXstrFldrPath = "" Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXstrFldrPathOUT = myXstrFldrPath
    Set myXobjFldrOUT = myXobjFldr
    myXstrDirPathOUT = myXstrDirPath
    myXstrFldrNameOUT = myXstrFldrName

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
    
'//S:�t�H���_�p�X���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
        
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXstrFldrPath = False
    myXstrFldrPath = Empty: Set myXobjFldr = Nothing
    myXstrDirPath = Empty: myXstrFldrName = Empty
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
    
'    If myXlonDirSlctOptn < 1 Or myXlonDirSlctOptn > 4 Then myXlonDirSlctOptn = 2
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
        
    myXlonDirSlctOptn = 2
    'myXlonDirSlctOptn = 1 : �A�N�e�B�u�u�b�N�̐e�t�H���_�p�X���擾
    'myXlonDirSlctOptn = 2 : FileDialog�I�u�W�F�N�g���g�p���ăt�H���_�p�X���擾
    'myXlonDirSlctOptn = 3 : �őO�ʂ̃G�N�X�v���[���ɕ\������Ă���t�H���_�p�X���擾
    'myXlonDirSlctOptn = 4 : �N�����̃G�N�X�v���[����I�����Ă��̃A�h���X�o�[���擾
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    If myXbisCurDirON = True Then myXstrDfltFldrPath = CurDir
    
  Dim myXbisTmp As Boolean
    If myXstrDfltFldrPath = "" Or myXstrDfltFldrPath = "C" Or _
            myXstrDfltFldrPath = "1" Or myXstrDfltFldrPath = "2" Then
        myXbisTmp = PfncbisCheckFolderExist(myXstrDfltFldrPath)
        If myXbisTmp = False Then myXstrDfltFldrPath = "2"
    End If
    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
    
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
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables3()
    
    myXbisExplrAdrsMsgOptn = True
    'myXbisExplrAdrsMsgOptn = True  : �E�B���h�I���̊m�F���b�Z�[�W��\������
    'myXbisExplrAdrsMsgOptn = False : �E�B���h�I���̊m�F���b�Z�[�W��\�����Ȃ�
    
End Sub

'SnsP_�t�H���_�p�X���擾
Private Sub snsProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt As String _
        = "�_�C�A���O�{�b�N�X���\������܂��̂ŁA�t�H���_��I�����ĉ������B"
    
'//�t�H���_�p�X�̎擾���@�ŕ��򂵂ăp�X���擾
    Select Case myXlonDirSlctOptn
        Case 1
        '//�A�N�e�B�u�u�b�N�̐e�t�H���_���擾
            myXstrFldrPath = ActiveWorkbook.Path
            
        Case 2
        '//FileDialog�I�u�W�F�N�g���g�p���ăt�H���_��I��
            Call setControlVariables2
            MsgBox coXstrMsgBxPrmpt
            Call instCSlctFldrPath
            
        Case 3
        '//CExplrAdrs�C���X�^���X���g�p���ăt�H���_���擾
            Call setControlVariables3
            Call instCExplrAdrs
            
        Case 4
        '//CExplrAdrsSlct�C���X�^���X���g�p���ăt�H���_���擾
            Call instCExplrAdrsSlct
            
        Case Else
    End Select
    If myXstrFldrPath = "" Then GoTo ExitPath
    
'//�w��t�H���_�̑��݂��m�F
    If PfncbisCheckFolderExist(myXstrFldrPath) = False Then
        myXstrFldrPath = ""
        GoTo ExitPath
    End If
    
'//�w��t�H���_�̃I�u�W�F�N�g���擾
    Set myXobjFldr = PfncobjGetFolder(myXstrFldrPath)
    
'//�w��t�H���_�̃t�H���_�������擾
    Call PfixGetFolderNameInformationByFSO(myXstrDirPath, myXstrFldrName, myXstrFldrPath)
    
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

'��ClassProc��_�t�H���_��I�����Ă��̃p�X���擾����
Private Sub instCSlctFldrPath()
  Dim myXinsFldrPath As CSlctFldrPath: Set myXinsFldrPath = New CSlctFldrPath
    With myXinsFldrPath
    '//�N���X���ϐ��ւ̓���
        .letDfltFldrPath = myXstrDfltFldrPath
        .letIniView = myXlonIniView
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        myXstrFldrPath = .fncstrDirectoryPath
    End With
    Set myXinsFldrPath = Nothing
End Sub

'��ClassProc��_�N�����̃G�N�X�v���[���̃A�h���X�o�[���擾����
Private Sub instCExplrAdrs()
  Dim myXinsExplrAdrs As CExplrAdrs: Set myXinsExplrAdrs = New CExplrAdrs
    With myXinsExplrAdrs
        .letMsgOptn = myXbisExplrAdrsMsgOptn
        myXstrFldrPath = .fncstrExplorerAddress
    End With
    Set myXinsExplrAdrs = Nothing
End Sub

'��ClassProc��_�N�����̃G�N�X�v���[����I�����Ă��̃A�h���X�o�[���擾����
Private Sub instCExplrAdrsSlct()
  Dim myXinsExplrAdrsSlct As CExplrAdrsSlct: Set myXinsExplrAdrsSlct = New CExplrAdrsSlct
    With myXinsExplrAdrsSlct
        myXstrFldrPath = .fncstrExplorerAddress
    End With
    Set myXinsExplrAdrsSlct = Nothing
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

 '��^�e_�w��t�H���_�̃I�u�W�F�N�g���擾����
Private Function PfncobjGetFolder(ByVal myXstrDirPath As String) As Object
    Set PfncobjGetFolder = Nothing
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FolderExists(myXstrDirPath) = False Then Exit Function
        Set PfncobjGetFolder = .GetFolder(myXstrDirPath)
    End With
    Set myXobjFSO = Nothing
End Function

 '��^�o_�w��t�H���_�̃t�H���_�������擾����(FileSystemObject�g�p)
Private Sub PfixGetFolderNameInformationByFSO( _
            myXstrPrntPath As String, myXstrDirName As String, _
            ByVal myXstrDirPath As String)
    myXstrPrntPath = Empty: myXstrDirName = Empty
    If myXstrDirPath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        myXstrPrntPath = .GetParentFolderName(myXstrDirPath)    '�e�t�H���_�p�X
        myXstrDirName = .GetFolder(myXstrDirPath).Name          '�t�H���_��
    End With
    Set myXobjFSO = Nothing
End Sub

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
'    myXlonDirSlctOptn = 2
'    'myXlonDirSlctOptn = 1 : �A�N�e�B�u�u�b�N�̐e�t�H���_�p�X���擾
'    'myXlonDirSlctOptn = 2 : FileDialog�I�u�W�F�N�g���g�p���ăt�H���_�p�X���擾
'    'myXlonDirSlctOptn = 3 : �őO�ʂ̃G�N�X�v���[���ɕ\������Ă���t�H���_�p�X���擾
'    'myXlonDirSlctOptn = 4 : �N�����̃G�N�X�v���[����I�����Ă��̃A�h���X�o�[���擾
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    If myXbisCurDirON = True Then myXstrDfltFldrPath = CurDir
'  Dim myXstrTmpPath As String
'    If myXstrDfltFldrPath = "" Or myXstrDfltFldrPath = "C" Or _
'            myXstrDfltFldrPath = "1" Or myXstrDfltFldrPath = "2" Then
'        myXstrTmpPath = PfncbisCheckFolderExist(myXstrDfltFldrPath)
'        If myXstrTmpPath = False Then myXstrDfltFldrPath = "2"
'    End If
'    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
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
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables3()
'    myXbisExplrAdrsMsgOptn = True
'    'myXbisExplrAdrsMsgOptn = True  : �E�B���h�I���̊m�F���b�Z�[�W��\������
'    'myXbisExplrAdrsMsgOptn = False : �E�B���h�I���̊m�F���b�Z�[�W��\�����Ȃ�
'End Sub
'��ModuleProc��_�f�B���N�g����I�����Ă��̃p�X���擾����
Private Sub callxRefSlctFldrPath()
'  Dim myXlonDirSlctOptn As Long, _
'        myXstrDfltFldrPath As String, myXlonIniView As Long, _
'        myXbisExplrAdrsMsgOptn As Boolean
'  Dim myXstrFldrPath As String, myXobjFldr As Object, _
'        myXstrDirPath As String, myXstrFldrName As String
    Call xRefSlctFldrPath.callProc( _
            myXstrFldrPath, myXobjFldr, myXstrDirPath, myXstrFldrName, _
            myXlonDirSlctOptn, myXstrDfltFldrPath, myXlonIniView, myXbisExplrAdrsMsgOptn)
    Debug.Print "�f�[�^: " & myXstrFldrPath
    Debug.Print "�f�[�^: " & myXstrDirPath
    Debug.Print "�f�[�^: " & myXstrFldrName
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSlctFldrPath()
'//xRefSlctFldrPath���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSlctFldrPath.resetConstant
End Sub
