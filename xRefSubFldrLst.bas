Attribute VB_Name = "xRefSubFldrLst"
'Includes CSubFldrLst
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSubFldrLst"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonFldrCnt As Long, myZobjFldr() As Object, _
            myZstrFldrName() As String, myZstrFldrPath() As String
    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
    'myZstrFldrName(k) : �t�H���_��
    'myZstrFldrPath(k) : �t�H���_�p�X
  
'//���͐���M��
  Private myXbisNotOutFldrInfo As Boolean
    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
  
'//���̓f�[�^
  Private myXstrDirPath As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXobjDir As Object

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjDir = Nothing
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
    Call callxRefSubFldrLst
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFldrCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            ByVal myXbisNotOutFldrInfoIN As Boolean, _
            ByVal myXstrDirPathIN As String)
    
'//���͕ϐ���������
    myXbisNotOutFldrInfo = False
    myXstrDirPath = Empty

'//���͕ϐ�����荞��
    myXbisNotOutFldrInfo = myXbisNotOutFldrInfoIN
    myXstrDirPath = myXstrDirPathIN
    
'//�o�͕ϐ���������
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT: Erase myZstrFldrNameOUT: Erase myZstrFldrPathOUT
    
'//�������s
    Call ctrProc
    If myXlonFldrCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    
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
    
'//S:�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾
    Call instCSubFldrLst
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
    myXlonFldrCnt = Empty
    Erase myZobjFldr: Erase myZstrFldrName: Erase myZstrFldrPath
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

    myXbisNotOutFldrInfo = False
    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
    
    myXstrDirPath = ActiveWorkbook.Path

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

'��ClassProc��_�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾����
Private Sub instCSubFldrLst()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSubFldrLst As CSubFldrLst: Set myXinsSubFldrLst = New CSubFldrLst
    With myXinsSubFldrLst
    '//�N���X���ϐ��ւ̓���
        .letNotOutFldrInfo = myXbisNotOutFldrInfo
        .letDirPath = myXstrDirPath
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonFldrCnt = .getFldrCnt
        If myXlonFldrCnt <= 0 Then GoTo JumpPath
        k = myXlonFldrCnt + Lo - 1
        ReDim myZobjFldr(k) As Object
        ReDim myZstrFldrName(k) As String
        ReDim myZstrFldrPath(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonFldrCnt
            Set myZobjFldr(k + Lo - 1) = .getFldrAry(k + Lc - 1)
            myZstrFldrName(k + Lo - 1) = .getFldrNameAry(k + Lc - 1)
            myZstrFldrPath(k + Lo - 1) = .getFldrPathAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSubFldrLst = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSubFldrLst = Nothing
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

'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾����
Private Sub callxRefSubFldrLst()
'  Dim myXbisNotOutFldrInfo As Boolean, myXstrDirPath As String
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String
'    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
'    'myZstrFldrName(k) : �t�H���_��
'    'myZstrFldrPath(k) : �t�H���_�p�X
    Call xRefSubFldrLst.callProc( _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, _
            myXbisNotOutFldrInfo, myXstrDirPath)
    Debug.Print "�f�[�^: " & myXlonFldrCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrLst()
'//xRefSubFldrLst���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSubFldrLst.resetConstant
End Sub
