Attribute VB_Name = "xRefSrchShtCmnt"
'Includes CSrchShtCmnt
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̃f�[�^�ƃR�����g���當������������ăf�[�^�ƈʒu�����擾����
'Rev.004
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSrchShtCmnt"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
    'myZstrTrgtVal(i) : �擾������
    'myZobjTrgtRng(i) : �s��ʒu�̃Z��
  
'//���̓f�[�^
  Private myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
            myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
            myXbisInStrOptn As Boolean
    'myZvarSrchCndtn(i, 1) : ����������
    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
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

'-----------------------------------------------------------------------------------------------

'PublicP_
Public Sub exeProc()
    
'//�������s
    Call callxRefSrchShtCmnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonTrgtValCntOUT As Long, _
            myZstrTrgtValOUT() As String, myZobjTrgtRngOUT() As Object, _
            ByVal myXlonSrchShtNoIN As Long, ByVal myXobjSrchSheetIN As Object, _
            ByVal myXlonShtSrchCntIN As Long, ByRef myZvarSrchCndtnIN As Variant, _
            ByVal myXbisInStrOptnIN As Boolean)
    
'//���͕ϐ���������
    myXlonSrchShtNo = Empty: Set myXobjSrchSheet = Nothing
    myXlonShtSrchCnt = Empty: myZvarSrchCndtn = Empty
    myXbisInStrOptn = False

'//���͕ϐ�����荞��
    myXlonSrchShtNo = myXlonSrchShtNoIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    myXlonShtSrchCnt = myXlonShtSrchCntIN
    myZvarSrchCndtn = myZvarSrchCndtnIN
    myXbisInStrOptn = myXbisInStrOptnIN
    
'//�o�͕ϐ���������
    myXlonTrgtValCntOUT = Empty
    Erase myZstrTrgtValOUT: Erase myZobjTrgtRngOUT
    
'//�������s
    Call ctrProc
    If myXlonTrgtValCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonTrgtValCntOUT = myXlonTrgtValCnt
    myZstrTrgtValOUT() = myZstrTrgtVal()
    myZobjTrgtRngOUT() = myZobjTrgtRng()

End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//�V�[�g��̃f�[�^���當������������ăf�[�^�ƈʒu�����擾
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXlonTrgtValCnt = Empty: Erase myZstrTrgtVal: Erase myZobjTrgtRng
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
    
    myXlonSrchShtNo = 2
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'    Set myXobjSrchSheet = ActiveSheet
    
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 3
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : ����������
    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
  Dim i As Long: i = L - 1
    i = i + 1   'i = 1
    myZvarSrchCndtn(i, L + 0) = "�e�t�H���_�p�X�F"
    myZvarSrchCndtn(i, L + 1) = 0
    myZvarSrchCndtn(i, L + 2) = 1
    myZvarSrchCndtn(i, L + 3) = 0
    i = i + 1   'i = 2
    myZvarSrchCndtn(i, L + 0) = "��������t�@�C���g���q�F"
    myZvarSrchCndtn(i, L + 1) = 0
    myZvarSrchCndtn(i, L + 2) = 1
    myZvarSrchCndtn(i, L + 3) = 0
    i = i + 1   'i = 3
    myZvarSrchCndtn(i, L + 0) = "�T�u�t�@�C���ꗗ"
    myZvarSrchCndtn(i, L + 1) = 1
    myZvarSrchCndtn(i, L + 2) = 0
    myZvarSrchCndtn(i, L + 3) = 0
    
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������

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

'��ClassProc��_�V�[�g��̃f�[�^���當������������ăf�[�^�ƈʒu�����擾����
Private Sub instCSrchShtCmnt()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSrchShtCmnt As CSrchShtCmnt: Set myXinsSrchShtCmnt = New CSrchShtCmnt
    With myXinsSrchShtCmnt
    '//�����񌟍��V�[�g�ƌ���������ݒ�
        Set .setSrchSheet = myXobjSrchSheet
        .letSrchCndtn = myZvarSrchCndtn
        .letInStrOptn = myXbisInStrOptn
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonTrgtValCnt = .getValCnt
        If myXlonTrgtValCnt <= 0 Then GoTo JumpPath
        i = myXlonTrgtValCnt + Lo - 1: j = Lo + 1
        ReDim myZstrTrgtVal(i) As String
        ReDim myZobjTrgtRng(i) As Object
        Lc = .getOptnBase
        For i = 1 To myXlonTrgtValCnt
            myZstrTrgtVal(i + Lo - 1) = .getValAry(i + Lc - 1)
            Set myZobjTrgtRng(i + Lo - 1) = .getPstnRngAry(i + Lc - 1)
        Next i
    End With
JumpPath:
    Set myXinsSrchShtCmnt = Nothing
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
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
''    Set myXobjSrchSheet = ActiveSheet
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'    myXlonShtSrchCnt = 3
'    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
'    'myZvarSrchCndtn(i, 1) : ����������
'    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
'    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
'    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
'  Dim k As Long: k = L - 1
'    k = k + 1   'k = 1
'    myZvarSrchCndtn(k, L + 0) = "�e�t�H���_�p�X�F"
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    k = k + 1   'k = 2
'    myZvarSrchCndtn(k, L + 0) = "��������t�@�C���g���q�F"
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    k = k + 1   'k = 3
'    myZvarSrchCndtn(k, L + 0) = "�T�u�t�@�C���ꗗ"
'    myZvarSrchCndtn(k, L + 1) = 1
'    myZvarSrchCndtn(k, L + 2) = 0
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'End Sub
'��ModuleProc��_�V�[�g��̃f�[�^�ƃR�����g���當������������ăf�[�^�ƈʒu�����擾����
Private Sub callxRefSrchShtCmnt()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'  Dim myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
'    'myZstrTrgtVal(i) : �擾������
'    'myZobjTrgtRng(i) : �s��ʒu�̃Z��
    Call xRefSrchShtCmnt.callProc( _
            myXlonTrgtValCnt, myZstrTrgtVal, myZobjTrgtRng, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn)
    If myXlonTrgtValCnt <= 0 Then Exit Sub
    Debug.Print "�f�[�^: " & myZstrTrgtVal(1)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefSrchShtCmnt()
'//xRefSrchShtCmnt���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSrchShtCmnt.resetConstant
End Sub
