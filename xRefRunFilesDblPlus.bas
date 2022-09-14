Attribute VB_Name = "xRefRunFilesDblPlus"
'Includes m1MexeFileLstup
'Includes m1MexeFileLstupSub
'Includes xRefRunFilesDbl
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�����t�@�C�������X�g�A�b�v���ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunFilesDblPlus"
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
  Dim myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
  Dim myXlonSubFileCnt As Long, myZstrSubFilePath() As String
    'myZstrOrgFilePath(i) : ���t�@�C���p�X1
    'myZstrSubFilePath(i) : ���t�@�C���p�X2
    

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
    myXlonSubFileCnt = Empty: Erase myZstrSubFilePath
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
    '����:  '��ModuleProc��_�����t�@�C�������X�g�A�b�v����
            '��ModuleProc��_2��ނ̕����t�@�C���ɑ΂��ĘA�����������{����
    '�o��: -
    
'//�������s
    Call callxRefRunFilesDblPlus
    
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
    
'//S:�����t�@�C�������X�g�A�b�v1
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:�����t�@�C�������X�g�A�b�v2
    Call snsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:2��ނ̕����t�@�C���ɑ΂��ĘA�����������{
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
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
End Sub

'SnsP_�����t�@�C�������X�g�A�b�v����1
Private Sub snsProc1()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object, myXstrExtsn As String
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
    
    Call m1MexeFileLstup.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonOrgFileCnt = myXlonFileCnt
    myZstrOrgFilePath() = myZstrFilePath()
    
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_�����t�@�C�������X�g�A�b�v����1
Private Sub snsProc2()
    myXbisExitFlag = False
    
  Dim myXbisCompFlag As Boolean
  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object, myXstrExtsn As String
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
    
    Call m1MexeFileLstupSub.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
    myXlonSubFileCnt = myXlonFileCnt
    myZstrSubFilePath() = myZstrFilePath()
    
    Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing: Set myXobjDirPstdCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_2��ނ̕����t�@�C���ɑ΂��ĘA�����������{����
Private Sub prsProc()
    myXbisExitFlag = False
  
  Dim myXbisCompFlag As Boolean
  Dim myXlonExeOrgFileCnt As Long, _
        myZstrExeOrgFileName() As String, myZstrOrgExeFilePath() As String
    'myZstrExeOrgFileName(i) : ���s�t�@�C����1
    'myZstrOrgExeFilePath(i) : ���s�t�@�C���p�X1
  Dim myXlonExeSubFileCnt As Long, _
        myZstrExeSubFileName() As String, myZstrSubExeFilePath() As String
    'myZstrExeSubFileName(i) : ���s�t�@�C����2
    'myZstrSubExeFilePath(i) : ���s�t�@�C���p�X2
    
    Call xRefRunFilesDbl.callProc( _
            myXbisCompFlag, _
            myXlonExeOrgFileCnt, myZstrExeOrgFileName, myZstrExeOrgFilePath, _
            myXlonExeSubFileCnt, myZstrExeSubFileName, myZstrExeSubFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath, _
            myXlonSubFileCnt, myZstrSubFilePath)
    If myXlonExeOrgFileCnt <= 0 Or myXlonExeSubFileCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrExeOrgFileName: Erase myZstrOrgExeFilePath
    Erase myZstrExeSubFileName: Erase myZstrSubExeFilePath
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
Private Sub callxRefRunFilesDblPlus()
  Dim myXbisCompFlag As Boolean
    Call xRefRunFilesDblPlus.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesDblPlus()
'//xRefRunFilesDblPlus���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunFilesDblPlus.resetConstant
End Sub
