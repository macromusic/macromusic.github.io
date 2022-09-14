Attribute VB_Name = "xRefFileDplct"
'Includes CFileDplct
'Includes PfnclonArrayDimension
'Includes PfncbisCheckFileExist
'Includes PfncbisCheckFolderExist
'Includes PfixGetFileNameInformationByFSO
'Includes PfncstrFilePathReplaceParentBaseExtension
'Includes PfncbisCheckArrayDimension
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�t�@�C���𕡐�����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefFileDplct"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  
'//���͐���M��
  
'//���̓f�[�^
  Private myXstrOrgFilePath As String
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : �����
  Private myXstrSaveDirPath As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonRunInfoNo As Long, myXstrRunInfo As String
  Private myXlonDplctFileCnt As Long, myZstrDplctFilePath() As String

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRunInfoNo = Empty: myXstrRunInfo = Empty
    myXlonDplctFileCnt = Empty: Erase myZstrDplctFilePath
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

'    myXlonOrgInfoCnt = 1
'    ReDim myZstrOrgInfo(myXlonOrgInfoCnt) As String
    
'//�������s
    Call callxRefFileDplct
    
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            ByVal myXstrOrgFilePathIN As String, _
            ByVal myXlonOrgInfoCntIN As Long, ByRef myZstrOrgInfoIN() As String, _
            ByVal myXstrSaveDirPathIN As String)
    
'//���͕ϐ���������
    myXstrOrgFilePath = Empty
    myXlonOrgInfoCnt = Empty
    Erase myZstrOrgInfo
    myXstrSaveDirPath = Empty

'//���͕ϐ�����荞��
    myXstrOrgFilePath = myXstrOrgFilePathIN
    If myXlonOrgInfoCntIN <= 0 Then Exit Sub
    myXlonOrgInfoCnt = myXlonOrgInfoCntIN
    myZstrOrgInfo() = myZstrOrgInfoIN()
    myXstrSaveDirPath = myXstrSaveDirPathIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
ExitPath:
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"    'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//�w��t�@�C���̑��݂��m�F
    If PfncbisCheckFileExist(myXstrOrgFilePath) = False Then GoTo ExitPath
    
'//�w��t�H���_�̑��݂��m�F
    If PfncbisCheckFolderExist(myXstrSaveDirPath) = False Then GoTo ExitPath
    
'//�w��t�@�C���̃t�@�C���������擾(FileSystemObject�g�p)
  Dim myXstrPrntPath As String, myXstrFileName As String, _
        myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformationByFSO( _
            myXstrPrntPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrOrgFilePath)
    
'//C:���ꗗ���������s
  Dim myXstrOrgPrnt As String, myXstrOrgBase As String, myXstrOrgExt As String
  Dim myXstrNewPrnt As String, myXstrNewBase As String, myXstrNewExt As String
  Dim myXstrDplctFilePath As String
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
        
    '//�t�@�C���p�X���̃f�B���N�g���p�Xor�x�[�X��or�g���q��u��
        myXstrOrgPrnt = Empty: myXstrNewPrnt = Empty
        myXstrOrgBase = Empty: myXstrNewBase = Empty
        
        myXstrOrgPrnt = myXstrPrntPath
        myXstrNewPrnt = myXstrSaveDirPath
        myXstrDplctFilePath = PfncstrFilePathReplaceParentBaseExtension( _
                                myXstrOrgFilePath, _
                                myXstrOrgPrnt, myXstrOrgBase, myXstrOrgExt, _
                                myXstrNewPrnt, myXstrNewBase, myXstrNewExt)
        myXstrOrgPrnt = Empty: myXstrNewPrnt = Empty
        
        myXstrOrgBase = myXstrBaseName
        myXstrNewBase = myXstrRunInfo
        myXstrDplctFilePath = PfncstrFilePathReplaceParentBaseExtension( _
                                myXstrDplctFilePath, _
                                myXstrOrgPrnt, myXstrOrgBase, myXstrOrgExt, _
                                myXstrNewPrnt, myXstrNewBase, myXstrNewExt)
        
        n = n + 1
        ReDim Preserve myZstrDplctFilePath(n) As String
        myZstrDplctFilePath(n) = myXstrDplctFilePath
NextPath:
    Next k
    myXlonDplctFileCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "7"    'PassFlag

'//�t�@�C���𕡐�
    Call instCFileDplct
    If myXbisExitFlag = True Then GoTo ExitPath
    
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

'RemP_�ۑ������ϐ������o��
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
    
'//�z��ϐ��̎��������擾
  Dim myXlonAryDmnsn As Long
    myXlonAryDmnsn = PfnclonArrayDimension(myZstrOrgInfo)
    
  Dim Li As Long, myXstrTmp As String
  
    On Error GoTo ExitPath
    Select Case myXlonAryDmnsn
        Case 1
            Li = LBound(myZstrOrgInfo): myXstrTmp = myZstrOrgInfo(Li)
            Exit Sub
        Case 2
            Li = LBound(myZstrOrgInfo, 1): myXstrTmp = myZstrOrgInfo(Li, Li)
        Case Else: GoTo ExitPath
    End Select
    On Error GoTo 0

  Dim myZstrOrgInfoINT() As String
    myZstrOrgInfoINT() = myZstrOrgInfo()
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
    i = UBound(myZstrOrgInfoINT, 1)
    ReDim myZstrOrgInfo(i) As String
    For i = LBound(myZstrOrgInfoINT) To UBound(myZstrOrgInfoINT)
        myZstrOrgInfo(i + Lo - Li) = myZstrOrgInfoINT(i)
    Next i
    
    Erase myZstrOrgInfoINT
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'CtrlP_
Private Sub ctrRunFiles()

'//C:���ꗗ���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim myXvarTmpInfo As Variant, k As Long: k = Li - 1
    For Each myXvarTmpPath In myZstrOrgInfo
        myXstrRunInfo = Empty
        k = k + 1: myXlonRunInfoNo = k
        myXstrRunInfo = CStr(myXvarTmpInfo)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next myXvarTmpInfo
    myXlonExeInfoCnt = n - Lo + 1
    
'//C:���ꗗ���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next k
    myXlonExeInfoCnt = n - Lo + 1
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
    
    myXstrOrgFilePath = ThisWorkbook.Path & "\�V�K �e�L�X�g �h�L�������g.txt"
    
    myXstrSaveDirPath = ThisWorkbook.Path & "\try"
    
End Sub

'checkP_�o�͕ϐ����e���m�F����
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_�g�p�����ϐ���ۑ�����
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

'��ClassProc��_�t�@�C���𕡐�����
Private Sub instCFileDplct()
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsFileDplct As CFileDplct: Set myXinsFileDplct = New CFileDplct
    With myXinsFileDplct
    '//�N���X���ϐ��ւ̓���
        .letOrgFilePath = myXstrOrgFilePath
        .letAutoNaming = True
        If myXlonDplctFileCnt <= 0 Then GoTo JumpPath
        .letDplctFileCnt = myXlonDplctFileCnt
        For i = 1 To myXlonDplctFileCnt
            .letDplctFilePathAry(i) = myZstrDplctFilePath(i + L - 1)
        Next i
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXbisExitFlag = Not .getCmpltFlag
    End With
    Set myXinsFileDplct = Nothing
JumpPath:
End Sub

'===============================================================================================
 
 '��^�e_�z��ϐ��̎��������擾����
Private Function PfnclonArrayDimension(ByRef myZvarOrgData As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarOrgData) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
End Function

 '��^�e_�w��t�@�C���̑��݂��m�F����
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '��^�e_�w��t�H���_�̑��݂��m�F����
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '��^�o_�w��t�@�C���̃t�@�C���������擾����(FileSystemObject�g�p)
Private Sub PfixGetFileNameInformationByFSO( _
            myXstrPrntPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String, _
            ByVal myXstrFilePath As String)
    myXstrPrntPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    If myXstrFilePath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExist As Boolean
    With myXobjFSO
        myXbisFileExist = .FileExists(myXstrFilePath)
        If myXbisFileExist = False Then Exit Sub
        myXstrPrntPath = .GetParentFolderName(myXstrFilePath)   '�e�t�H���_�p�X
        myXstrFileName = .GetFileName(myXstrFilePath)           '�t�@�C����
        myXstrBaseName = .GetBaseName(myXstrFilePath)           '�t�@�C���x�[�X��
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         '�t�@�C���g���q
    End With
    Set myXobjFSO = Nothing
End Sub

 '��^�e_�t�@�C���p�X���̃f�B���N�g���p�Xor�x�[�X��or�g���q��u������
Private Function PfncstrFilePathReplaceParentBaseExtension( _
            ByVal myXstrOrgFilePath As String, _
            ByVal myXstrOrgPrnt As String, ByVal myXstrOrgBase As String, _
            ByVal myXstrOrgExt As String, _
            ByVal myXstrNewPrnt As String, ByVal myXstrNewBase As String, _
            ByVal myXstrNewExt As String) As String
    PfncstrFilePathReplaceParentBaseExtension = Empty
    If myXstrOrgFilePath = "" Then Exit Function
    If myXstrNewPrnt = "" And myXstrNewBase = "" And myXstrNewExt = "" Then Exit Function
  Dim myXstrNewFilePath As String
    If InStr(myXstrOrgFilePath, myXstrOrgPrnt) > 0 And myXstrNewPrnt <> "" Then _
        myXstrNewFilePath = Replace(myXstrOrgFilePath, myXstrOrgPrnt, myXstrNewPrnt)
    If InStr(myXstrOrgFilePath, myXstrOrgBase) > 0 And myXstrNewBase <> "" Then _
        myXstrNewFilePath = Replace(myXstrOrgFilePath, myXstrOrgBase, myXstrNewBase)
    If InStr(myXstrOrgFilePath, myXstrOrgExt) > 0 And myXstrNewExt <> "" Then _
        myXstrNewFilePath = Replace(myXstrOrgFilePath, myXstrOrgExt, myXstrNewExt)
    PfncstrFilePathReplaceParentBaseExtension = myXstrNewFilePath
End Function

 '��^�e_�z��ϐ��̎��������w�莟���ƈ�v���邩���`�F�b�N����
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarOrgData As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarOrgData) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
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

'��ModuleProc��_�t�@�C���𕡐�����
Private Sub callxRefFileDplct()
  Dim myXstrOrgFilePath As String
  Dim myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
  Dim myXstrSaveDirPath As String
    'myZstrOrgInfo(i) : �����
    myXstrOrgFilePath = ThisWorkbook.Path & "\�V�K �e�L�X�g �h�L�������g.txt"
    myXlonOrgInfoCnt = 2
    ReDim myZstrOrgInfo(2) As String
    myZstrOrgInfo(1) = "A"
    myZstrOrgInfo(2) = "B"
    myXstrSaveDirPath = ThisWorkbook.Path & "\try"
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonExeInfoCnt As Long, myZstrExeInfo() As String
'    'myZstrExeInfo(i) : ���s���
    Call xRefFileDplct.callProc( _
            myXbisCmpltFlag, _
            myXstrOrgFilePath, myXlonOrgInfoCnt, myZstrOrgInfo, myXstrSaveDirPath)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub PforResetConstantInxRefFileDplct()
'//xRefFileDplct���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefFileDplct.resetConstant
End Sub
