Attribute VB_Name = "xRefSlctFilePathRptExtd"
'Includes CSlctFilePathRpt
'Includes CVrblToSht
'Includes PincPickUpExtensionMatchFilePathArray
'Includes PfncbisCheckFileExtension
'Includes PfixGetFileFor1DArray
'Includes PfixGetFolderFileStringInformationFor1DArray
'Includes PfncstrCheckAndGetFilesParentFolder
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�w�蕶������܂ރt�@�C�����̃t�@�C�����J�Ԃ��I�����Ă��̃p�X���擾���ăV�[�g�ɏ����o��
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSlctFilePathRptExtd"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
            myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(i) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(i) : �t�@�C����
    'myZstrFilePath(i) : �t�@�C���p�X
  Private myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
  
'//���͐���M��
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��
  
'//���̓f�[�^
  Private myXstrDfltFldrPath As String
    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
  Private myXstrDfltFilePath As String
    'myXstrDfltFilePath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFilePath = "1" : ���̃u�b�N���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFilePath = "2" : �A�N�e�B�u�u�b�N���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFilePath = "*" : �f�t�H���g�p�X���w��
  Private myXstrExtsn As String
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
  Private myXlonOrdrCnt As Long, myXlonTrgtWrdCnt As Long, myZvarOdrTrgtWrdPos() As Variant
    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i�Ԗڂ̒��o�t�@�C���̎w�蕶����:����p
    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : �w�蕶������x�[�X�t�@�C�����̐擪�Ɋ܂�
    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : �w�蕶������x�[�X�t�@�C�����̐ڔ��Ɋ܂�
    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : �w�蕶������x�[�X�t�@�C�������Ɋ܂�
  Private myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean, myXbisPstFlag As Boolean
  Private myXbisCurDirON As Boolean
    'myXbisCurDirON = False : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肵�Ȃ�
    'myXbisCurDirON = True  : �f�t�H���g�p�X�ɃJ�����g�f�B���N�g����ݒ肷��

'//���W���[�����ϐ�_�f�[�^
  Private myXlonFileOrgCnt As Long, myZstrFilePathOrg() As String
  Private myZvarPstData As Variant, myXobjPstFrstCell As Object
  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myXobjPstdCell As Object
    
'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    myXbisCurDirON = False
    
    myXlonFileOrgCnt = Empty: Erase myZstrFilePathOrg
    myZvarPstData = Empty: Set myXobjPstFrstCell = Nothing
    myXbisInptBxOFF = False: myXbisEachWrtON = False
    Set myXobjPstdCell = Nothing
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
    Call callxRefSlctFilePathRptExtd
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, myXobjDirPstdCellOUT As Object, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrDfltFldrPathIN As String, ByVal myXstrDfltFilePathIN As String, _
            ByVal myXstrExtsnIN As String, _
            ByRef myZstrAddFltrIN() As String, _
            ByVal myXbisFltrClrIN As Boolean, ByVal myXlonFltrIndxIN As Long, _
            ByVal myXlonIniViewIN As Long, ByVal myXbisMultSlctIN As Boolean, _
            ByVal myXlonOrdrCntIN As Long, ByVal myXlonTrgtWrdCntIN As Long, _
            ByRef myZvarOdrTrgtWrdPosIN() As Variant, _
            ByVal myXobjDirPstFrstCellIN As Object, ByVal myXobjFilePstFrstCellIN As Object)
    
'//���͕ϐ���������
    myXlonOutputOptn = Empty
    
    myXstrDfltFldrPath = Empty: myXstrDfltFilePath = Empty
    myXstrExtsn = Empty
    Erase myZstrAddFltr: myXbisFltrClr = False: myXlonFltrIndx = Empty
    myXlonIniView = Empty: myXbisMultSlct = False
    myXlonOrdrCnt = Empty: myXlonTrgtWrdCnt = Empty: Erase myZvarOdrTrgtWrdPos
    Set myXobjDirPstFrstCell = Nothing: Set myXobjFilePstFrstCell = Nothing
    
'//���͕ϐ�����荞��
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrDfltFldrPath = myXstrDfltFldrPathIN
    myXstrDfltFilePath = myXstrDfltFilePathIN
    myXstrExtsn = myXstrExtsnIN
    myZstrAddFltr() = myZstrAddFltrIN()
    myXbisFltrClr = myXbisFltrClrIN
    myXlonFltrIndx = myXlonFltrIndxIN
    myXlonIniView = myXlonIniViewIN
    myXbisMultSlct = myXbisMultSlctIN
    myXlonOrdrCnt = myXlonOrdrCntIN
    myXlonTrgtWrdCnt = myXlonTrgtWrdCntIN
    myZvarOdrTrgtWrdPos() = myZvarOdrTrgtWrdPosIN()
    Set myXobjDirPstFrstCell = myXobjDirPstFrstCellIN
    Set myXobjFilePstFrstCell = myXobjFilePstFrstCellIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    myXlonFileCntOUT = Empty: Erase myZobjFileOUT
    Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
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
    Call setControlVariables1
    Call setControlVariables2
    
'//S:�t�@�C���p�X���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"  'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"  'PassFlag
    
'//Run:�t�@�C���p�X���V�[�g�ɏ����o��
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
    myXbisCmpltFlag = False
    myXlonFileCnt = Empty: Erase myZobjFile
    Erase myZstrFileName: Erase myZstrFilePath
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
Private Sub setControlVariables1()
        
    myXstrDfltFldrPath = "1"
    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
    
    myXstrDfltFilePath = "1"
    'myXstrDfltFilePath = ""  : �f�t�H���g�p�X�w�薳��
    'myXstrDfltFilePath = "1" : ���̃u�b�N���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFilePath = "2" : �A�N�e�B�u�u�b�N���f�t�H���g�p�X�Ɏw��
    'myXstrDfltFilePath = "*" : �f�t�H���g�p�X���w��
    
    myXstrExtsn = ""
    
    ReDim myZstrAddFltr(1, 2) As String
    myZstrAddFltr(1, 1) = "PDF�t�@�C��"
    myZstrAddFltr(1, 2) = "*.pdf"
    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
    
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
    
    myXlonOrdrCnt = 2
    myXlonTrgtWrdCnt = 2
    ReDim myZvarOdrTrgtWrdPos(myXlonOrdrCnt, myXlonTrgtWrdCnt, 2) As Variant
    myZvarOdrTrgtWrdPos(1, 1, 1) = "C"
    myZvarOdrTrgtWrdPos(1, 1, 2) = 1
    myZvarOdrTrgtWrdPos(1, 2, 1) = "Mtch"
    myZvarOdrTrgtWrdPos(1, 2, 2) = 2
    myZvarOdrTrgtWrdPos(2, 1, 1) = "C"
    myZvarOdrTrgtWrdPos(2, 1, 2) = 1
    myZvarOdrTrgtWrdPos(2, 2, 1) = "Sort"
    myZvarOdrTrgtWrdPos(2, 2, 2) = 2
    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i�Ԗڂ̒��o�t�@�C���̎w�蕶����:����p
    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : �w�蕶������x�[�X�t�@�C�����̐擪�Ɋ܂�
    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : �w�蕶������x�[�X�t�@�C�����̐ڔ��Ɋ܂�
    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : �w�蕶������x�[�X�t�@�C�������Ɋ܂�
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    myXlonOutputOptn = 3
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��

'    myZvarVrbl = 1
    
'    Set myXobjDirPstFrstCell = Selection
'    Set myXobjFilePstFrstCell = Selection
    
    myXbisInptBxOFF = False
    'myXbisInptBxOFF = False : �w��ʒu�������ꍇ��InputBox�Ŕ͈͎w�肷��
    'myXbisInptBxOFF = True  : �w��ʒu�������ꍇ��InputBox�Ŕ͈͎w�肵�Ȃ�
    
    myXbisEachWrtON = False
    'myXbisEachWrtON = False : �z��ϐ����f�[�^����x�ɏ����o������
    'myXbisEachWrtON = True  : �z��ϐ����f�[�^��1�f�[�^�Â����o������
    
End Sub

'SnsP_�t�@�C���p�X���擾
Private Sub snsProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt As String _
        = "�_�C�A���O�{�b�N�X���\������܂��̂ŁA�t�@�C����I�����ĉ������B"
    
'//�t�@�C����I�����Ă��̃p�X���擾
    MsgBox coXstrMsgBxPrmpt
    Call instCSlctFilePathRpt
    If myXlonFileOrgCnt <= 0 Then GoTo ExitPath
    
'//�擾����2�����z��f�[�^��1�����z��f�[�^�ɓ���ւ���
  Dim myXlonTmpCnt As Long, myZstrTmp() As String
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim myXvarTmp As Variant, n As Long: n = L - 1
    For Each myXvarTmp In myZstrFilePathOrg
        If myXvarTmp = "" Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrTmp(n) As String
        myZstrTmp(n) = myXvarTmp
NextPath:
    Next myXvarTmp
    
'//�擾�����t�@�C���p�X�ꗗ����g���q�őI��
  Dim myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String
    Call PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt, myZstrExtMtchFilePath, _
            myZstrTmp, myXstrExtsn)
    If myXlonExtMtchFileCnt <= 0 Then GoTo ExitPath
    
'//�t�@�C���p�X�ꗗ����t�@�C���I�u�W�F�N�g�ꗗ���擾
    Call PfixGetFileFor1DArray(myXlonFileCnt, myZobjFile, myZstrExtMtchFilePath)

'//�t�@�C���ꗗ�̃t�@�C�������擾
  Dim myXlonInfoCnt As Long
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFileName, _
            myZobjFile, 1)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath

'//�t�@�C���ꗗ�̃t�@�C���p�X���擾
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFilePath, _
            myZobjFile, 2)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExtMtchFilePath
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
    If myXlonOutputOptn = 2 Then
    '//�t�@�C�����������o���ꍇ
        myZvarPstData = myZstrFileName
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt2
        
    ElseIf myXlonOutputOptn = 3 Then
    '//�e�t�H���_�ɉ����ď����o���ꍇ
    
    '//�f�B���N�g���p�X���G�N�Z���V�[�g�ɏ����o��
        myZvarPstData = myXstrPrntPath
        Set myXobjPstFrstCell = myXobjDirPstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt3
        
        Call instCVrblToSht
        If myXbisPstFlag = False Then GoTo ExitPath
        Set myXobjDirPstdCell = myXobjPstdCell
        
    '//�t�@�C�������G�N�Z���V�[�g�ɏ����o��
        myZvarPstData = myZstrFileName
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt2
        
    Else
    '//�t�@�C���p�X�������o���ꍇ
        myZvarPstData = myZstrFilePath
        Set myXobjPstFrstCell = myXobjFilePstFrstCell
        
        If myXbisInptBxOFF = False And myXobjPstFrstCell Is Nothing Then _
            MsgBox coXstrMsgBxPrmpt1
            
    End If
    
'//�t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    Call instCVrblToSht
    If myXbisPstFlag = False Then GoTo ExitPath
    Set myXobjFilePstdCell = myXobjPstdCell
    
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

'��ClassProc��_�w�蕶������܂ރt�@�C�����̃t�@�C�����J�Ԃ��I�����Ă��̃p�X���擾����
Private Sub instCSlctFilePathRpt()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsFilePathRpt As CSlctFilePathRpt: Set myXinsFilePathRpt = New CSlctFilePathRpt
    With myXinsFilePathRpt
    '//�N���X���ϐ��ւ̓���
        .letDfltFldrPath = myXstrDfltFldrPath
        .letDfltFilePath = myXstrDfltFilePath
        .letExtsn = myXstrExtsn
        .letAddFltr = myZstrAddFltr
        .letFltrClr = myXbisFltrClr
        .letFltrIndx = myXlonFltrIndx
        .letIniView = myXlonIniView
        .letMultSlct = myXbisMultSlct
        .letOdrTrgtWrdPosAry = myZvarOdrTrgtWrdPos
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonFileOrgCnt = .getFileCnt
        If myXlonFileOrgCnt <= 0 Then GoTo JumpPath
        k = myXlonFileOrgCnt + Lo - 1
        ReDim myZstrFilePathOrg(k, Lo) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFileOrgCnt
            myZstrFilePathOrg(k + Lo - 1, Lo) = .getFilePathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsFilePathRpt = Nothing
    Call variablesOfCSlctFilePathRpt(myXlonFileOrgCnt, myZstrFilePathOrg)    'Debug.Print
End Sub
Private Sub variablesOfCSlctFilePathRpt( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//CSlctFilePathRpt�N���X������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim i As Long, j As Long
    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
            Debug.Print "�f�[�^" & i & "," & j & ": " & myZvarField(i, j)
        Next j
    Next i
End Sub

'��ClassProc��_�ϐ������G�N�Z���V�[�g�ɏ����o��
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//�N���X���ϐ��ւ̓���
        .letVrbl = myZvarPstData
        Set .setPstFrstCell = myXobjPstFrstCell
        .letInptBxOFF = myXbisInptBxOFF
        .letEachWrtON = myXbisEachWrtON
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        myXbisPstFlag = .fncbisCmpltFlag
        Set myXobjPstdCell = .getPstdRng
    End With
    Set myXinsVrblToSht = Nothing
End Sub

'===============================================================================================

 '��^�o_�t�@�C���ꗗ����w��g���q�ƈ�v����t�@�C���p�X�𒊏o����
Private Sub PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String, _
            ByRef myXstrOrgFilePath() As String, ByVal myXstrExtsn As String)
'Includes PfncbisCheckFileExtension
'myZstrExtMtchFilePath(i) : ���o�t�@�C���p�X
'myXstrOrgFilePath(i) : ���t�@�C���p�X
    myXlonExtMtchFileCnt = Empty: Erase myZstrExtMtchFilePath
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myXstrOrgFilePath): myXstrTmp = myXstrOrgFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarFilePath As Variant, myXbisExtChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarFilePath In myXstrOrgFilePath
      Dim myXstrFilePath As String: myXstrFilePath = myXvarFilePath
        myXbisExtChck = PfncbisCheckFileExtension(myXstrFilePath, myXstrExtsn)
        If myXbisExtChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExtMtchFilePath(n) As String
        myZstrExtMtchFilePath(n) = myXvarFilePath
NextPath:
    Next
    myXlonExtMtchFileCnt = n - Lo + 1
    myXvarFilePath = Empty
ExitPath:
End Sub

 '��^�e_�w��t�@�C�����w��g���q�ł��邱�Ƃ��m�F����
Private Function PfncbisCheckFileExtension( _
            ByVal myXstrFilePath As String, ByVal myXstrExtsn As String) As Boolean
'myXstrExtsn = "*" : �C�ӂ̕�����̃��C���h�J�[�h
    PfncbisCheckFileExtension = False
    If myXstrFilePath = "" Then Exit Function
    If myXstrExtsn = "" Then GoTo JumpPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrOrgExt As String
    With myXobjFSO
        If .FileExists(myXstrFilePath) = False Then Exit Function
        myXstrOrgExt = .GetExtensionName(myXstrFilePath)
    End With
  Dim myXstrDesExt As String: myXstrDesExt = myXstrExtsn
    If Left(myXstrDesExt, 1) = "." Then myXstrDesExt = Mid(myXstrDesExt, 2)
    myXstrOrgExt = LCase(myXstrOrgExt)
    myXstrDesExt = LCase(myXstrDesExt)
    If myXstrOrgExt = myXstrDesExt Then GoTo JumpPath
  Dim myXlonPstn As Long: myXlonPstn = InStr(myXstrDesExt, "*")
    Select Case myXlonPstn
        Case 1
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
        Case Len(myXstrExtsn)
            If Left(myXstrOrgExt, Len(myXstrDesExt) - 1) _
                    <> Left(myXstrDesExt, Len(myXstrDesExt) - 1) Then _
                Exit Function
        Case Else
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
            If Left(myXstrOrgExt, myXlonPstn - 1) _
                    <> Left(myXstrDesExt, myXlonPstn - 1) Then _
                Exit Function
    End Select
    Set myXobjFSO = Nothing
JumpPath:
    PfncbisCheckFileExtension = True
End Function

 '��^�o_1�����z��̃t�@�C���p�X����t�@�C���I�u�W�F�N�g�ꗗ���擾����
Private Sub PfixGetFileFor1DArray( _
                myXlonFileCnt As Long, myZobjFile() As Object, _
                ByRef myZstrFilePath() As String)
'myZobjFile(i) : �t�@�C���I�u�W�F�N�g�ꗗ
'myZstrFilePath(i) : ���t�@�C���p�X�ꗗ
    myXlonFileCnt = Empty: Erase myZobjFile
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFilePath): myXstrTmp = myZstrFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjTmp As Object, i As Long, n As Long: n = Lo - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath) To UBound(myZstrFilePath)
        myXstrTmp = Empty
        myXstrTmp = myZstrFilePath(i)
        With myXobjFSO
            If .FileExists(myXstrTmp) = False Then GoTo NextPath
            Set myXobjTmp = .GetFile(myXstrTmp)
        End With
        n = n + 1: ReDim Preserve myZobjFile(n) As Object
        Set myZobjFile(n) = myXobjTmp
NextPath:
    Next i
    myXlonFileCnt = n - Lo + 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '��^�o_1�����z��̃t�H���_�t�@�C���I�u�W�F�N�g�ꗗ�̕���������擾����
Private Sub PfixGetFolderFileStringInformationFor1DArray( _
                myXlonInfoCnt As Long, myZstrInfo() As String, _
                ByRef myZobjFldrFile() As Object, _
                Optional ByVal coXlonStrOptn As Long = 1)
'myZstrInfo(i) : ���o�t�H���_���
'myZobjFldrFile(i) : ���t�H���_or���t�@�C��
'coXlonStrOptn = 1  : ���O (Name)
'coXlonStrOptn = 2  : �p�X (Path)
'coXlonStrOptn = 3  : �e�t�H���_ (ParentFolder)
'coXlonStrOptn = 4  : ���� (Attributes)
'coXlonStrOptn = 5  : ��� (Type)
    myXlonInfoCnt = Empty: Erase myZstrInfo
  Dim myXobjTmp As Object, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXstrTmp As String, i As Long, n As Long: n = Lo - 1
    On Error GoTo NextPath
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        myXstrTmp = Empty
        If myZobjFldrFile(i) Is Nothing Then GoTo NextPath
        Select Case coXlonStrOptn
            Case 1: myXstrTmp = myZobjFldrFile(i).Name
            Case 2: myXstrTmp = myZobjFldrFile(i).Path
            Case 3: myXstrTmp = myZobjFldrFile(i).ParentFolder
            Case 4: myXstrTmp = myZobjFldrFile(i).Attributes
            Case 5: myXstrTmp = myZobjFldrFile(i).Type
        End Select
        n = n + 1: ReDim Preserve myZstrInfo(n) As String
        myZstrInfo(n) = myXstrTmp
NextPath:
    Next i
    On Error GoTo 0
    myXlonInfoCnt = n - Lo + 1
ExitPath:
End Sub

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
'Private Sub setControlVariables()
'    myXstrDfltFldrPath = "1"
'    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
'    myXstrDfltFilePath = "1"
'    'myXstrDfltFilePath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFilePath = "1" : ���̃u�b�N���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFilePath = "2" : �A�N�e�B�u�u�b�N���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFilePath = "*" : �f�t�H���g�p�X���w��
'    myXstrExtsn = ""
'    ReDim myZstrAddFltr(1, 2) As String
'    myZstrAddFltr(1, 1) = "PDF�t�@�C��"
'    myZstrAddFltr(1, 2) = "*.pdf"
'    'myZstrAddFltr(i, 1) : �t�@�C���̌����w�肷�镶����(�t�@�C��)
'    'myZstrAddFltr(i, 2) : �t�@�C���̌����w�肷�镶����(�t�B���^������)
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
'    myXlonOrdrCnt = 2
'    myXlonTrgtWrdCnt = 2
'    ReDim myZvarOdrTrgtWrdPos(myXlonOrdrCnt, myXlonTrgtWrdCnt, 2) As Variant
'    myZvarOdrTrgtWrdPos(1, 1, 1) = "C"
'    myZvarOdrTrgtWrdPos(1, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(1, 2, 1) = "Mtch"
'    myZvarOdrTrgtWrdPos(1, 2, 2) = 2
'    myZvarOdrTrgtWrdPos(2, 1, 1) = "C"
'    myZvarOdrTrgtWrdPos(2, 1, 2) = 1
'    myZvarOdrTrgtWrdPos(2, 2, 1) = "Sort"
'    myZvarOdrTrgtWrdPos(2, 2, 2) = 2
'    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i�Ԗڂ̒��o�t�@�C���̎w�蕶����:����p
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : �w�蕶������x�[�X�t�@�C�����̐擪�Ɋ܂�
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : �w�蕶������x�[�X�t�@�C�����̐ڔ��Ɋ܂�
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : �w�蕶������x�[�X�t�@�C�������Ɋ܂�
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXlonOutputOptn = 3
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��
''    myZvarVrbl = 1
''    Set myXobjDirPstFrstCell = Selection
''    Set myXobjFilePstFrstCell = Selection
'    myXbisInptBxOFF = False
'    'myXbisInptBxOFF = False : �w��ʒu�������ꍇ��InputBox�Ŕ͈͎w�肷��
'    'myXbisInptBxOFF = True  : �w��ʒu�������ꍇ��InputBox�Ŕ͈͎w�肵�Ȃ�
'    myXbisEachWrtON = False
'    'myXbisEachWrtON = False : �z��ϐ����f�[�^����x�ɏ����o������
'    'myXbisEachWrtON = True  : �z��ϐ����f�[�^��1�f�[�^�Â����o������
'End Sub
'��ModuleProc��_�w�蕶������܂ރt�@�C�����̃t�@�C�����J�Ԃ��I�����Ă��̃p�X���擾���ăV�[�g�ɏ����o��
Private Sub callxRefSlctFilePathRptExtd()
'  Dim myXlonOutputOptn As Long, _
'        myXstrDfltFldrPath As String, myXstrDfltFilePath As String, myXstrExtsn As String, _
'        myZstrAddFltr() As String, myXbisFltrClr As Boolean, myXlonFltrIndx As Long, _
'        myXlonIniView As Long, myXbisMultSlct As Boolean, _
'        myXlonOrdrCnt As Long, myXlonTrgtWrdCnt As Long, myZvarOdrTrgtWrdPos() As Variant, _
'        myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�H���_�p�X�^�����G�N�Z���V�[�g�ɏ����o��
'    'myXstrDfltFldrPath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFldrPath = "C" : C�h���C�u���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "1" : ���̃u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "2" : �A�N�e�B�u�u�b�N�̐e�t�H���_���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFldrPath = "*" : �f�t�H���g�p�X���w��
'    'myXstrDfltFilePath = ""  : �f�t�H���g�p�X�w�薳��
'    'myXstrDfltFilePath = "1" : ���̃u�b�N���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFilePath = "2" : �A�N�e�B�u�u�b�N���f�t�H���g�p�X�Ɏw��
'    'myXstrDfltFilePath = "*" : �f�t�H���g�p�X���w��
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
'    'myZvarOdrTrgtWrdPos(i, p, 1) = x : i�Ԗڂ̒��o�t�@�C���̎w�蕶����:����p
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 1 : �w�蕶������x�[�X�t�@�C�����̐擪�Ɋ܂�
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 2 : �w�蕶������x�[�X�t�@�C�����̐ڔ��Ɋ܂�
'    'myZvarOdrTrgtWrdPos(i, p, 2) = 3 : �w�蕶������x�[�X�t�@�C�������Ɋ܂�
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
'    'myZobjFile(i) : �t�@�C���I�u�W�F�N�g
'    'myZstrFileName(i) : �t�@�C����
'    'myZstrFilePath(i) : �t�@�C���p�X
    Call xRefSlctFilePathRptExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXlonOutputOptn, _
            myXstrDfltFldrPath, myXstrDfltFilePath, myXstrExtsn, _
            myZstrAddFltr, myXbisFltrClr, myXlonFltrIndx, _
            myXlonIniView, myXbisMultSlct, _
            myXlonOrdrCnt, myXlonTrgtWrdCnt, myZvarOdrTrgtWrdPos, _
            myXobjDirPstFrstCell, myXobjFilePstFrstCell)
    Debug.Print "�f�[�^: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSlctFilePathRptExtd()
'//xRefSlctFilePathRptExtd���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSlctFilePathRptExtd.resetConstant
End Sub
