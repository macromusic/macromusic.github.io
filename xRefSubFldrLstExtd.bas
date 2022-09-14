Attribute VB_Name = "xRefSubFldrLstExtd"
'Includes CSubFldrLst
'Includes CVrblToSht
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾���ăV�[�g�ɏ����o��
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSubFldrLstExtd"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonFldrCnt As Long, myZobjFldr() As Object, _
            myZstrFldrName() As String, myZstrFldrPath() As String
    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
    'myZstrFldrName(k) : �t�H���_��
    'myZstrFldrPath(k) : �t�H���_�p�X
  Private myXobjPstdCell As Object
  
'//���͐���M��
  Private myXbisNotOutFldrInfo As Boolean
    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
  
'//���̓f�[�^
  Private myXstrDirPath As String
  
  Private myXlonOutputOptn As Long
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
  
  Private myXobjFldrPstFrstCell As Object
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXobjDir As Object

  Private myXbisInptBxOFF As Boolean, myXbisEachWrtON As Boolean
  Private myZvarPstVrbl As Variant
    
'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjDir = Nothing
    myXbisInptBxOFF = False: myXbisEachWrtON = False
    myZvarPstVrbl = Empty
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
    Call callxRefSubFldrLstExtd
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            myXobjPstdCellOUT As Object, _
            ByVal myXbisNotOutFldrInfoIN As Boolean, _
            ByVal myXstrDirPathIN As String, _
            ByVal myXlonOutputOptnIN As Long, ByVal myXobjFldrPstFrstCellIN As Object)
    
'//���͕ϐ���������
    myXbisNotOutFldrInfo = False
    myXstrDirPath = Empty
    myXlonOutputOptn = Empty
    Set myXobjFldrPstFrstCell = Nothing

'//���͕ϐ�����荞��
    myXbisNotOutFldrInfo = myXbisNotOutFldrInfoIN
    myXstrDirPath = myXstrDirPathIN
    myXlonOutputOptn = myXlonOutputOptnIN
    Set myXobjFldrPstFrstCell = myXobjFldrPstFrstCellIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT: Erase myZstrFldrNameOUT: Erase myZstrFldrPathOUT
    Set myXobjPstdCellOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    Set myXobjPstdCellOUT = myXobjPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariablesA
    Call setControlVariablesB
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾
    Call instCSubFldrLst
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:�擾�f�[�^�����H
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:�ϐ������G�N�Z���V�[�g�ɏ����o��
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
'
'    If myXlonOutputOptn < 0 Or myXlonOutputOptn > 2 Then myXlonDirSlctOptn = 1
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesA()

    myXbisNotOutFldrInfo = False
    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
    
'    myXstrDirPath = ActiveWorkbook.Path
    myXstrDirPath = "C:\Users\Hiroki\Documents\_VBA4XPC\11 �v���O�����f�[�^�x�[�X\02_VBA���W���[��"

End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariablesB()
    
    myXlonOutputOptn = 1
    'myXlonOutputOptn = 0 : �����o����������
    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��

'    myZvarVrbl = 1
    
'    Set myXobjFldrPstFrstCell = Selection
    
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

'PrcsP_�擾�f�[�^�����H
Private Sub prsProc()
    myXbisExitFlag = False
    
    On Error GoTo ExitPath
    Select Case myXlonOutputOptn
    '//�t�H���_�p�X��I��
        Case 1: myZvarPstVrbl = myZstrFldrPath
        
    '//�t�H���_����I��
        Case 2: myZvarPstVrbl = myZstrFldrName
        
        Case Else: Exit Sub
    End Select
    On Error GoTo 0
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�ϐ������G�N�Z���V�[�g�ɏ����o��
Private Sub runProc()
    myXbisExitFlag = False
  Const coXstrMsgBxPrmpt1 As String _
        = "�t�H���_�p�X��\��t����ʒu���w�肵�ĉ������B"
  Const coXstrMsgBxPrmpt2 As String _
        = "�t�H���_����\��t����ʒu���w�肵�ĉ������B"
    
'//�ϐ����������o�����ŕ���
    Select Case myXlonOutputOptn
    '//�G�N�Z���V�[�g�ɏ����o��
        Case 1
            If myXbisInptBxOFF = False And myXobjFldrPstFrstCell Is Nothing Then _
                MsgBox coXstrMsgBxPrmpt1
            Call instCVrblToSht
        
    '//�G�N�Z���V�[�g�ɏ����o��
        Case 2
            If myXbisInptBxOFF = False And myXobjFldrPstFrstCell Is Nothing Then _
                MsgBox coXstrMsgBxPrmpt2
            Call instCVrblToSht
        
        Case Else: Exit Sub
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
    
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

'��ClassProc��_�ϐ������G�N�Z���V�[�g�ɏ����o��
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//�N���X���ϐ��ւ̓���
        .letVrbl = myZvarPstVrbl
        Set .setPstFrstCell = myXobjFldrPstFrstCell
        .letInptBxOFF = myXbisInptBxOFF
        .letEachWrtON = myXbisEachWrtON
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXbisExitFlag = Not .getCmpltFlag
        Set myXobjPstdCell = .getPstdRng
    End With
    Set myXinsVrblToSht = Nothing
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
'Private Sub setControlVariablesA()
'    myXbisNotOutFldrInfo = False
'    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
'    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
'    myXstrDirPath = ActiveWorkbook.Path
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariablesB()
'    myXlonOutputOptn = 1
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �t�H���_�p�X���G�N�Z���V�[�g�ɏ����o��
'    'myXlonOutputOptn = 2 : �t�H���_�����G�N�Z���V�[�g�ɏ����o��
''    myZvarVrbl = 1
''    Set myXobjFldrPstFrstCell = Selection
'End Sub
'��ModuleProc��_�w��f�B���N�g�����̃T�u�t�H���_�ꗗ���擾���ăV�[�g�ɏ����o��
Private Sub callxRefSubFldrLstExtd()
'  Dim myXbisNotOutFldrInfo As Boolean, myXstrDirPath As String, _
'        myXlonOutputOptn As Long, myXobjFldrPstFrstCell As Object
'    'myXbisNotOutFldrInfo = False : �t�H���_�I�u�W�F�N�ƃt�H���_���𗼕��o�͂���
'    'myXbisNotOutFldrInfo = True  : �t�H���_�I�u�W�F�N�g�̂ݏo�͂��ăt�H���_���͏o�͂��Ȃ�
'    'myXlonOutputOptn = 0 : �����o����������
'    'myXlonOutputOptn = 1 : �G�N�Z���V�[�g�ɏ����o��
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String, myXobjPstdCell As Object
'    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
'    'myZstrFldrName(k) : �t�H���_��
'    'myZstrFldrPath(k) : �t�H���_�p�X
    Call xRefSubFldrLstExtd.callProc( _
            myXbisCmpltFlag, _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, myXobjPstdCell, _
            myXbisNotOutFldrInfo, myXstrDirPath, myXlonOutputOptn, myXobjFldrPstFrstCell)
    Debug.Print "�f�[�^: " & myXlonFldrCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSubFldrLstExtd()
'//xRefSubFldrLstExtd���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSubFldrLstExtd.resetConstant
End Sub
