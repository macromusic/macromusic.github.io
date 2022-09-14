Attribute VB_Name = "m1MexeFileLstup"
'Includes m1Msub1ShtFileLst
'Includes m1Msub2SlctFldrPathExtd
'Includes m1Msub3SubFileLstExtd
'Includes PfncbisCheckFolderExist
'Includes PfixChangeModuleConstValue

Option Explicit
Option Base 1

'��ModuleProc��_�����t�@�C�������X�g�A�b�v����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "m1MexeFileLstup"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
        myZstrFileName() As String, myZstrFilePath() As String, _
        myXobjFilePstdCell As Object, _
        myXstrDirPath As String, myXobjDirPstdCell As Object, _
        myXstrExtsn As String
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Dim myXbisDirPstFlag As Boolean
    'myXbisDirPstFlag = True  : �e�t�H���_�p�X�̓\��t���L��
    'myXbisDirPstFlag = False : �e�t�H���_�p�X�̓\��t������
  
'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisDirPstFlag = False
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
    '����:  '��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X�ꗗ���擾����
            '��ModuleProc��_�t�H���_��I�����Ă��̃p�X���擾���ăV�[�g�ɏ����o��
            '��ModuleProc��_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ăV�[�g�ɏ����o��
    '�o��: -

    
'//�������s
    Call callm1MexeFileLstup
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdCellOUT As Object, _
            myXstrDirPathOUT As String, myXobjDirPstdCellOUT As Object, _
            myXstrExtsnOUT As String)
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT
    Erase myZstrFileNameOUT
    Erase myZstrFilePathOUT
    Set myXobjFilePstdCellOUT = Nothing
    myXstrDirPathOUT = Empty
    Set myXobjDirPstdCellOUT = Nothing
    myXstrExtsnOUT = Empty
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonFileCnt <= 0 Then Exit Sub
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    Set myXobjFilePstdCellOUT = myXobjFilePstdCell
    
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell
    myXstrExtsnOUT = myXstrExtsn
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//S:�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X�ꗗ���擾
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:�t�@�C����I�����Ă��̃p�X���擾���ăV�[�g�ɏ����o��
    Call snsProc2
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//S:�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ăV�[�g�ɏ����o��
    Call snsProc3
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonFileCnt = Empty: Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdCell = Nothing
    myXstrDirPath = Empty: Set myXobjDirPstdCell = Nothing
    myXstrExtsn = Empty
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

'SnsP_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X�ꗗ���擾����
Private Sub snsProc1()
    myXbisExitFlag = False
    
  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
        myXbisInStrOptn As Boolean
  Dim myXbisRowDrctn As Boolean
    
    Call m1Msub1ShtFileLst.callProc( _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn, myXbisRowDrctn)
    
    Set myXobjSrchSheet = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_�t�@�C����I�����Ă��̃p�X���擾���ăV�[�g�ɏ����o��
Private Sub snsProc2()
    myXbisExitFlag = False
    
    If myXlonFileCnt > 0 Then Exit Sub
    If PfncbisCheckFolderExist(myXstrDirPath) = True Then
        myXbisDirPstFlag = True
        Exit Sub
    End If

  Dim myXlonOutputOptn As Long, myXobjDirPstFrstCell As Object
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
    
    If myXobjDirPstdCell Is Nothing Then
        myXlonOutputOptn = 0
        Set myXobjDirPstFrstCell = Nothing
    Else
        myXlonOutputOptn = 1
        Set myXobjDirPstFrstCell = myXobjDirPstdCell
    End If
    
  Dim myXlonDirSlctOptn As Long, _
        myXstrDfltFldrPath As String, myXlonIniView As Long, _
        myXbisExplrAdrsMsgOptn As Boolean
  
  Dim myXstrFldrPath As String, myXobjFldr As Object, _
        myXstrPrntPath As String, myXstrFldrName As String
    Call m1Msub2SlctFldrPathExtd.callProc( _
            myXbisDirPstFlag, _
            myXstrFldrPath, myXobjFldr, myXstrPrntPath, myXstrFldrName, _
            myXobjDirPstdCell, _
            myXlonDirSlctOptn, myXstrDfltFldrPath, myXlonIniView, myXbisExplrAdrsMsgOptn, _
            myXlonOutputOptn, myXobjDirPstFrstCell)
'    Debug.Print "�f�[�^: " & myXstrFldrPath
'    Debug.Print "�f�[�^: " & myXstrPrntPath
'    Debug.Print "�f�[�^: " & myXstrFldrName
    
    If myXstrFldrPath = "" Then GoTo ExitPath
    myXstrDirPath = myXstrFldrPath
    
    Set myXobjDirPstFrstCell = Nothing
    Set myXobjFldr = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_�w��f�B���N�g�����̃T�u�t�@�C���ꗗ���擾���ăV�[�g�ɏ����o��
Private Sub snsProc3()
    myXbisExitFlag = False
    
    If myXlonFileCnt > 0 Then Exit Sub
    
  Dim myXbisNotOutFileInfo As Boolean, myXlonOutputOptn As Long, _
        myXlonSrchOptn As Long, _
        myXobjDirPstFrstCell As Object, myXobjFilePstFrstCell As Object
    
    Set myXobjDirPstFrstCell = myXobjDirPstdCell
    Set myXobjFilePstFrstCell = myXobjFilePstdCell
    
  Dim myXbisCompFlag As Boolean
    Call m1Msub3SubFileLstExtd.callProc( _
            myXbisCompFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, myXobjDirPstdCell, _
            myXbisNotOutFileInfo, myXlonOutputOptn, _
            myXstrDirPath, myXstrExtsn, myXlonSrchOptn, _
            myXobjDirPstFrstCell, myXobjFilePstFrstCell)
    Debug.Print "�f�[�^: " & myXlonFileCnt
    
    Set myXobjDirPstFrstCell = Nothing
    Set myXobjFilePstFrstCell = Nothing
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

 '��^�e_�w��t�H���_�̑��݂��m�F����
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
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

'��ModuleProc��_�����t�H���_�����X�g�A�b�v����
Private Sub callm1MexeFileLstup()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdCell As Object, _
'        myXstrDirPath As String, myXobjDirPstdCell As Object, myXstrExtsn As String
'    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
'    'myZstrFileName(k) : �t�@�C����
'    'myZstrFilePath(k) : �t�@�C���p�X
    Call m1MexeFileLstup.callProc( _
            myXbisCmpltFlag, _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn)
    Call variablesOfm1MexeFileLstup(myXlonFileCnt, myZstrFilePath)    'Debug.Print
End Sub
Private Sub variablesOfm1MexeFileLstup( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//m1MexeFileLstup������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myZvarField) To UBound(myZvarField)
        Debug.Print "�f�[�^" & k & ": " & myZvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInm1MexeFileLstup()
'//m1MexeFileLstup���W���[���̃��W���[���������̃��Z�b�g����
    Call m1MexeFileLstup.resetConstant
End Sub
