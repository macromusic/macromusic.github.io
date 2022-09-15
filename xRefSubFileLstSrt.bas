Attribute VB_Name = "xRefSubFileLstSrt"
'Includes CSubFldrFileLst
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ă\�[�g����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSubFileLstSrt"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
  
'//���͐���M��
  Private myXbisNotOutFileInfo As Boolean
    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�
  
'//���̓f�[�^
  Private myXstrDirPath As String, myXstrExtsn As String
  Private myXlonFileSortOptn As Long
    'myXlonFileSortOptn = 1 : �\�[�g���Ȃ�
    'myXlonFileSortOptn = 2 : �t�@�C�������Ƀ\�[�g����
    'myXlonFileSortOptn = 3 : �X�V�������Ƀ\�[�g����
  
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
    Call callxRefSubFileLstSrt
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            ByVal myXbisNotOutFileInfoIN As Boolean, _
            ByVal myXstrDirPathIN As String, ByVal myXstrExtsnIN As String, _
            ByVal myXlonFileSortOptnIN As Long)
    
'//���͕ϐ���������
    myXbisNotOutFileInfo = False
    
    myXstrDirPath = Empty: myXstrExtsn = Empty
    myXlonFileSortOptn = Empty

'//���͕ϐ�����荞��
    myXbisNotOutFileInfo = myXbisNotOutFileInfoIN
    
    myXstrDirPath = myXstrDirPathIN
    myXstrExtsn = myXstrExtsnIN
    myXlonFileSortOptn = myXlonFileSortOptnIN
    
'//�o�͕ϐ���������
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    
'//�������s
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    
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
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾
    Call instCSubFileLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
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
Private Sub setControlVariables()

    myXbisNotOutFileInfo = False
    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�
    
    myXstrDirPath = ActiveWorkbook.Path
    
    myXstrExtsn = ""

    myXlonFileSortOptn = 3
    'myXlonFileSortOptn = 1 : �\�[�g���Ȃ�
    'myXlonFileSortOptn = 2 : �t�@�C�������Ƀ\�[�g����
    'myXlonFileSortOptn = 3 : �X�V�������Ƀ\�[�g����
    
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

'��ClassProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾����
Private Sub instCSubFileLst()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSubFileLst As CSubFileLst: Set myXinsSubFileLst = New CSubFileLst
    With myXinsSubFileLst
    '//�N���X���ϐ��ւ̓���
        .letFileSortOptn = myXlonFileSortOptn
        .letNotOutFileInfo = myXbisNotOutFileInfo
        .letDirPath = myXstrDirPath
        .letExtsn = myXstrExtsn
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
    Set myXinsSubFileLst = Nothing
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
'    myXbisNotOutFileInfo = False
'    'myXbisNotOutFileInfo = False : �t�@�C���I�u�W�F�N�ƃt�@�C�����𗼕��o�͂���
'    'myXbisNotOutFileInfo = True  : �t�@�C���I�u�W�F�N�g�̂ݏo�͂��ăt�@�C�����͏o�͂��Ȃ�
'    myXstrDirPath = ActiveWorkbook.Path
'    myXstrExtsn = ""
'    myXlonFileSortOptn = 1
'    'myXlonFileSortOptn = 1 : �\�[�g���Ȃ�
'    'myXlonFileSortOptn = 2 : �t�@�C�������Ƀ\�[�g����
'    'myXlonFileSortOptn = 3 : �X�V�������Ƀ\�[�g����
'End Sub
'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ă\�[�g����
Private Sub callxRefSubFileLstSrt()
'  Dim myXbisNotOutFileInfo As Boolean, _
'        myXstrDirPath As String, myXstrExtsn As String, myXlonFileSortOptn As Long
'    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
'    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
'    'myXlonFileSortOptn = 1 : �\�[�g���Ȃ�
'    'myXlonFileSortOptn = 2 : �t�@�C�������Ƀ\�[�g����
'    'myXlonFileSortOptn = 3 : �X�V�������Ƀ\�[�g����
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String
'    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
'    'myZstrFileName(k) : �t�@�C����
'    'myZstrFilePath(k) : �t�@�C���p�X
    Call xRefSubFileLstSrt.callProc( _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXbisNotOutFileInfo, myXstrDirPath, myXstrExtsn, myXlonFileSortOptn)
    Debug.Print "�f�[�^: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFileLstSrt()
'//xRefSubFileLstSrt���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSubFileLstSrt.resetConstant
End Sub
