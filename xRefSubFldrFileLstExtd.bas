Attribute VB_Name = "xRefSubFldrFileLstExtd"
'Includes CSubFldrFileLst
'Includes CVrblToSht
'Includes PfncstrCheckAndGetFilesParentFolder
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ăV�[�g�ɏ����o��
'Rev.003
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSubFldrFileLstExtd"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
  Private myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
  
'//���͐���M��
  Private myXbisNotOutFileInfo As Boolean
    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��
  
'//���̓f�[�^
  Private myXstrDirPath As String, myXstrExtsn As String
    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�
  Private myXlonSrchOptn As Long
    'myXlonSrchOptn = 1 : �w��t�H���_�����̃t�@�C���̃p�X�̂ݎ擾
    'myXlonSrchOptn = 2 : �w��t�H���_�����̃t�@�C���ƃT�u�t�H���_���̃t�@�C���̃p�X���擾
    'myXlonSrchOptn = 3 : �w��t�H���_�����̃T�u�t�H���_���̃t�@�C���̃p�X�̂ݎ擾
  Private myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean, myXbisPstFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonFileOrgCnt As Long, myZstrFilePathOrg() As String
  Private myZvarPstData As Variant, myXobjPstFrstCell As Object
  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myXobjPstdCell As Object

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False: myXbisPstFlag = False
    
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
    Call callxRefSubFldrFileLstExtd
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, myXobjDirPstdCellOUT As Object, _
            ByVal myXbisNotOutFileInfoIN As Boolean, _
            ByVal myXlonOutputOptnIN As Long, _
            ByVal myXstrDirPathIN As String, ByVal myXstrExtsnIN As String, _
            ByVal myXlonSrchOptnIN As Long, _
            ByVal myXobjDirPstFrstCellIN As Object, ByVal myXobjFilePstFrstCellIN As Object)
    
'//���͕ϐ���������
    myXbisNotOutFileInfo = False
    myXlonOutputOptn = Empty
    
    myXstrDirPath = Empty: myXstrExtsn = Empty
    myXlonSrchOptn = Empty
    Set myXobjDirPstFrstCell = Nothing: Set myXobjFilePstFrstCell = Nothing

'//���͕ϐ�����荞��
    myXbisNotOutFileInfo = myXbisNotOutFileInfoIN
    myXlonOutputOptn = myXlonOutputOptnIN
    
    myXstrDirPath = myXstrDirPathIN
    myXstrExtsn = myXstrExtsnIN
    myXlonSrchOptn = myXlonSrchOptnIN
    Set myXobjDirPstFrstCell = myXobjDirPstFrstCellIN
    Set myXobjFilePstFrstCell = myXobjFilePstFrstCellIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    
'//�������s
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
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
    
'//S:�w��f�B���N�g�����̕����T�u�t�H���_���̃T�u�t�@�C���ꗗ���擾
    Call instCSubFldrFileLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:�t�@�C���p�X���V�[�g�ɏ����o��
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
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
    
'    If myXstrDirPath = "" Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()

    myXstrDirPath = ActiveWorkbook.Path
    
    myXstrExtsn = ""
    
    myXlonSrchOptn = 1
    'myXlonSrchOptn = 1 : �w��t�H���_�����̃t�@�C���̃p�X�̂ݎ擾
    'myXlonSrchOptn = 2 : �w��t�H���_�����̃t�@�C���ƃT�u�t�H���_���̃t�@�C���̃p�X���擾
    'myXlonSrchOptn = 3 : �w��t�H���_�����̃T�u�t�H���_���̃t�@�C���̃p�X�̂ݎ擾
    
    myXbisNotOutFileInfo = False
    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�

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

'SnsP_
Private Sub snsProc()
    myXbisExitFlag = False
    
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

'��ClassProc��_�w��f�B���N�g�����̕����T�u�t�H���_���̃T�u�t�@�C���ꗗ���擾����
Private Sub instCSubFldrFileLst()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSubFldrFileLst As CSubFldrFileLst
    Set myXinsSubFldrFileLst = New CSubFldrFileLst
    With myXinsSubFldrFileLst
    '//�N���X���ϐ��ւ̓���
        .letDirPath = myXstrDirPath
        .letExtsn = myXstrExtsn
        .letSrchOptn = myXlonSrchOptn
        .letNotOutFileInfo = myXbisNotOutFileInfo
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonFileCnt = .getFileCnt
        If myXlonFileCnt <= 0 Then GoTo JumpPath
        k = myXlonFileCnt + Lo - 1
        ReDim myZobjFile(k) As Object
        ReDim myZstrFileName(k) As String
        ReDim myZstrFilePath(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFileCnt
            Set myZobjFile(k + Lo - 1) = .getFileAry(k + Lc - 1)
            myZstrFileName(k + Lo - 1) = .getFileNameAry(k + Lc - 1)
            myZstrFilePath(k + Lo - 1) = .getFilePathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSubFldrFileLst = Nothing
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
'    myXstrDirPath = ActiveWorkbook.Path
'    myXstrExtsn = ""
'    myXlonSrchOptn = 1
'    'myXlonSrchOptn = 1 : �w��t�H���_�����̃t�@�C���̃p�X�̂ݎ擾
'    'myXlonSrchOptn = 2 : �w��t�H���_�����̃t�@�C���ƃT�u�t�H���_���̃t�@�C���̃p�X���擾
'    'myXlonSrchOptn = 3 : �w��t�H���_�����̃T�u�t�H���_���̃t�@�C���̃p�X�̂ݎ擾
'    myXbisNotOutFileInfo = False
'    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
'    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�
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
'End Sub
'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ăV�[�g�ɏ����o��
Private Sub callxRefSubFldrFileLstExtd()
'  Dim myXbisNotOutFileInfo As Boolean, myXlonOutputOptn As Long, _
'        myXstrDirPath As String, myXstrExtsn As String, myXlonSrchOptn As Long, _
'        myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
'    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
'    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�@�C���p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�@�C�������G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 3 : �e�t�H���_�ɉ����ăt�@�C���p�X�^�����G�N�Z���V�[�g�ɏ����o��
'    'myXlonSrchOptn = 1 : �w��t�H���_�����̃t�@�C���̃p�X�̂ݎ擾
'    'myXlonSrchOptn = 2 : �w��t�H���_�����̃t�@�C���ƃT�u�t�H���_���̃t�@�C���̃p�X���擾
'    'myXlonSrchOptn = 3 : �w��t�H���_�����̃T�u�t�H���_���̃t�@�C���̃p�X�̂ݎ擾
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, myXobjDirPstdCell As Object
'    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
'    'myZstrFileName(k) : �t�@�C����
'    'myZstrFilePath(k) : �t�@�C���p�X
    Call xRefSubFldrFileLstExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXbisNotOutFileInfo, myXlonOutputOptn, _
            myXstrDirPath, myXstrExtsn, myXlonSrchOptn, _
            myXobjDirPstFrstCell, myXobjFilePstFrstCell)
    Debug.Print "�f�[�^: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrFileLstExtd()
'//xRefSubFldrFileLstExtd���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSubFldrFileLstExtd.resetConstant
End Sub
