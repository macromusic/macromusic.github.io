Attribute VB_Name = "xRefVrblToTxt"
'Includes CVrblToTxt
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�ϐ������e�L�X�g�t�@�C���ɏ����o��
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefVrblToTxt"
  Private Const meMlonExeNum As Long = 0
 
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//���̓f�[�^
  Private myZvarVrbl As Variant, myXstrSpltChar As String, myXstrSaveFilePath As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
  Private myXbisMsgBoxON As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myZvarPstVrbl As Variant

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXbisMsgBoxON = False
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
    Call callxRefVrblToTxt
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            ByVal myZvarVrblIN As Variant, _
            ByVal myXstrSpltCharIN As String, ByVal myXstrSaveFilePathIN As String)
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
    
'//���͕ϐ���������
    myZvarVrbl = Empty
    myXstrSpltChar = Empty
    myXstrSaveFilePath = Empty

'//���͕ϐ�����荞��
    myZvarVrbl = myZvarVrblIN
    myXstrSpltChar = myXstrSpltCharIN
    myXstrSaveFilePath = myXstrSaveFilePathIN
    
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
    
'//S:
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:�ϐ������e�L�X�g�t�@�C���ɏ����o��
    Call instCVrblToTxt
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
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

    myXbisMsgBoxON = True
    'myXbisMsgBxON = False : �ϐ��̃e�L�X�g�����o��������MsgBox��\�����Ȃ�
    'myXbisMsgBxON = True  : �ϐ��̃e�L�X�g�����o��������MsgBox��\������
    
    ReDim myZvarVrbl(2, 2) As Variant
    myZvarVrbl(1, 1) = "A"
    myZvarVrbl(1, 2) = "A"
    myZvarVrbl(2, 1) = "A"
    myZvarVrbl(2, 2) = "A"
    
    myXstrSpltChar = ""
    
  Dim myXstrPrntPath As String, myXstrFileName As String
    myXstrPrntPath = ThisWorkbook.Path
    myXstrFileName = "testOUT.txt"
    myXstrSaveFilePath = myXstrPrntPath & "\" & myXstrFileName
    
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
    
    On Error GoTo ExitPath
    myZvarPstVrbl = myZvarVrbl
    On Error GoTo 0
    
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

'��ClassProc��_�ϐ������e�L�X�g�t�@�C���ɏ����o��
Private Sub instCVrblToTxt()
  Dim myXbisCmpltFlag As Boolean
  Dim myXinsVrblToTxt As CVrblToTxt: Set myXinsVrblToTxt = New CVrblToTxt
    With myXinsVrblToTxt
    '//�N���X���ϐ��ւ̓���
        .letVrbl = myZvarVrbl
        .letSpltChar = myXstrSpltChar
        .letSaveFilePath = myXstrSaveFilePath
        .letMsgBoxON = True
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        myXbisExitFlag = Not .fncbisCmpltFlag
    End With
    Set myXinsVrblToTxt = Nothing
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
'    myXbisMsgBoxON = True
'    'myXbisMsgBxON = False : �ϐ��̃e�L�X�g�����o��������MsgBox��\�����Ȃ�
'    'myXbisMsgBxON = True  : �ϐ��̃e�L�X�g�����o��������MsgBox��\������
'    ReDim myZvarVrbl(2, 2) As Variant
'    myZvarVrbl(1, 1) = "A"
'    myZvarVrbl(1, 2) = "A"
'    myZvarVrbl(2, 1) = "A"
'    myZvarVrbl(2, 2) = "A"
'    myXstrSpltChar = ""
'  Dim myXstrPrntPath As String, myXstrFileName As String
'    myXstrPrntPath = ThisWorkbook.Path
'    myXstrFileName = "testOUT.txt"
'    myXstrSaveFilePath = myXstrPrntPath & "\" & myXstrFileName
'End Sub
'��ModuleProc��_�ϐ������e�L�X�g�t�@�C���ɏ����o��
Private Sub callxRefVrblToTxt()
'  Dim myZvarVrbl As Variant, myXstrSpltChar As String, myXstrSaveFilePath As String
'  Dim myXbisCmpltFlag As Boolean
    Call xRefVrblToTxt.callProc(myXbisCmpltFlag, myZvarVrbl, myXstrSpltChar, myXstrSaveFilePath)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefVrblToTxt()
'//xRefVrblToTxt���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefVrblToTxt.resetConstant
End Sub
