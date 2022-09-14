Attribute VB_Name = "xRefXlsOpenClose"
'Includes CXlsOpen
'Includes CXlsClose
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckFileExist
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���u�b�N���J���ď�������
'Rev.002
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefXlsOpenClose"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXobjOpndBook As Object
  
'//���͐���M��
  
'//���̓f�[�^
  Private myXbisOpnRdOnly As Boolean
  Private myXstrOpnFullName As String
  
  Private myXbisSaveON As Boolean
  Private myXstrSaveBkName As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  Private myXbisErrFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXstrCloseFullName As String
  Private myXbisBkOpnd As Boolean
    'myXbisBkOpnd = True  : �w��G�N�Z���u�b�N���J���Ă���
    'myXbisBkOpnd = False : �w��G�N�Z���u�b�N���J���Ă��Ȃ�
  Private myXbisBkRdOnly As Boolean
    'myXbisBkRdOnly = True  : �w��G�N�Z���u�b�N���ǂݎ���p�ŊJ���Ă���
    'myXbisBkRdOnly = False : �w��G�N�Z���u�b�N���ǂݎ���p�ł͊J���Ă��Ȃ�

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrCloseFullName = Empty
    myXbisBkOpnd = False: myXbisBkRdOnly = False
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
    Call callxRefXlsOpenClose
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXobjOpndBookOUT As Object, _
            ByVal myXbisOpnRdOnlyIN As Boolean, ByVal myXstrOpnFullNameIN As String, _
            ByVal myXbisSaveONIN As Boolean, ByVal myXstrSaveBkNameIN As String)
    
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    
'//���͕ϐ���������
    myXbisOpnRdOnly = False
    myXstrOpnFullName = Empty
    
    myXbisSaveON = False
    myXstrSaveBkName = Empty

'//���͕ϐ�����荞��
    myXbisOpnRdOnly = myXbisOpnRdOnlyIN
    myXstrOpnFullName = myXstrOpnFullNameIN
    
    myXbisSaveON = myXbisSaveONIN
    myXstrSaveBkName = myXstrSaveBkNameIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    Set myXobjOpndBookOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    Set myXobjOpndBookOUT = myXobjOpndBook

ExitPath:
End Sub

'PublicF_
Public Function fncobjOpenedBook( _
            ByVal myXbisOpnRdOnlyIN As Boolean, ByVal myXstrOpnFullNameIN As String, _
            ByVal myXbisSaveONIN As Boolean, ByVal myXstrSaveBkNameIN As String) As Object
    Set fncobjOpenedBook = Nothing
    
'//���͕ϐ���������
    myXbisOpnRdOnly = False
    myXstrOpnFullName = Empty
    
    myXbisSaveON = False
    myXstrSaveBkName = Empty

'//���͕ϐ�����荞��
    myXbisOpnRdOnly = myXbisOpnRdOnlyIN
    myXstrOpnFullName = myXstrOpnFullNameIN
    
    myXbisSaveON = myXbisSaveONIN
    myXstrSaveBkName = myXstrSaveBkNameIN
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Function
    
    Set fncobjOpenedBook = myXobjOpndBook
    
End Function

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:
    Call setControlVariables1
    Call setControlVariables2
    
'//S:�G�N�Z���u�b�N���J��
    Call snsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc1
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:�G�N�Z���u�b�N�����
    Call prsProc2
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
    Set myXobjOpndBook = Nothing
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
    
'//�w��t�@�C���̑��݂��m�F
    If PfncbisCheckFileExist(myXstrOpnFullName) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    
    myXbisOpnRdOnly = True
    'myXbisOpnRdOnly = False : �w��G�N�Z���u�b�N��ǂݎ���p�ɂ����ɊJ��
    'myXbisOpnRdOnly = True  : �w��G�N�Z���u�b�N��ǂݎ���p�ŊJ��
    
    myXstrOpnFullName = ThisWorkbook.Path & "\�V�K Microsoft Excel ���[�N�V�[�g.xlsx"
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    myXbisSaveON = False
    'myXbisSaveON = False : ����O�ɕۑ����Ȃ�
    'myXbisSaveON = True  : ����O�ɕۑ�����
    
    myXstrSaveBkName = ""
    
End Sub

'SnsP_�G�N�Z���u�b�N���J��
Private Sub snsProc1()
    myXbisExitFlag = False

    Call instCXlsOpen
    If myXobjOpndBook Is Nothing Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_
Private Sub prsProc1()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3-1"   'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�G�N�Z���u�b�N�����
Private Sub prsProc2()
    myXbisExitFlag = False

    myXstrCloseFullName = myXstrOpnFullName
    
    Call instCXlsClose
    If myXbisErrFlag = True Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_
Private Sub runProc()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5-1"   'PassFlag
    
    'XarbProgCode
    
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

'���W���[�����o_
Private Sub MsubProc()
End Sub

'���W���[�����e_
Private Function MfncFunc() As Variant
End Function

'===============================================================================================

'��ClassProc��_�G�N�Z���u�b�N���J��
Private Sub instCXlsOpen()
  Dim myXinsXlsOpen As CXlsOpen: Set myXinsXlsOpen = New CXlsOpen
    With myXinsXlsOpen
    '//�N���X���ϐ��ւ̓���
        .letOpnRdOnly = myXbisOpnRdOnly
        .letOpnFullName = myXstrOpnFullName
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXbisCmpltFlag = .getCmpltFlag
        Set myXobjOpndBook = .getOpndBook
        Set myXobjOpndBook = .fncobjOpenedBook
    End With
    Set myXinsXlsOpen = Nothing
End Sub

'��ClassProc��_�G�N�Z���u�b�N�����
Private Sub instCXlsClose()
  Dim myXinsXlsClose As CXlsClose: Set myXinsXlsClose = New CXlsClose
    With myXinsXlsClose
    '//�N���X���ϐ��ւ̓���
        .letCloseFullName = myXstrCloseFullName
        .letSaveON = myXbisSaveON
        .letSaveBkName = myXstrSaveBkName
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXbisErrFlag = Not .getCmpltFlag
        myXbisBkOpnd = .getBkOpnd
        myXbisBkRdOnly = .getBkRdOnly
    End With
    Set myXinsXlsClose = Nothing
End Sub

'===============================================================================================

 '��^�o_
Private Sub PfixProc()
End Sub

 '��^�e_
Private Function PfncFunc() As Variant
End Function

 '��^�e_�z��ϐ��̎��������w�莟���ƈ�v���邩���`�F�b�N����
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarDataAry As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarDataAry) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarDataAry, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
End Function
 
 '��^�e_�w��t�@�C���̑��݂��m�F����
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
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
'    myXbisOpnRdOnly = True
'    'myXbisOpnRdOnly = False : �w��G�N�Z���u�b�N��ǂݎ���p�ɂ����ɊJ��
'    'myXbisOpnRdOnly = True  : �w��G�N�Z���u�b�N��ǂݎ���p�ŊJ��
'    myXstrOpnFullName = ""
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXbisSaveON = False
'    'myXbisSaveON = False : ����O�ɕۑ����Ȃ�
'    'myXbisSaveON = True  : ����O�ɕۑ�����
'    myXstrSaveBkName = ""
'End Sub
'��ModuleProc��_�G�N�Z���u�b�N���J���ď�������
Private Sub callxRefXlsOpenClose()
  Dim myXbisOpnRdOnly As Boolean, myXstrOpnFullName As String
    'myXbisOpnRdOnly = False : �w��G�N�Z���u�b�N��ǂݎ���p�ɂ����ɊJ��
    'myXbisOpnRdOnly = True  : �w��G�N�Z���u�b�N��ǂݎ���p�ŊJ��
    myXstrOpnFullName = ThisWorkbook.Path & "\�V�K Microsoft Excel ���[�N�V�[�g.xlsx"
  Dim myXbisSaveON As Boolean, myXstrSaveBkName As String
    'myXbisSaveON = False : ����O�ɕۑ����Ȃ�
    'myXbisSaveON = True  : ����O�ɕۑ�����
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjOpndBook As Object
    Call xRefXlsOpenClose.callProc( _
            myXbisCmpltFlag, myXobjOpndBook, _
            myXbisOpnRdOnly, myXstrOpnFullName, _
            myXbisSaveON, myXstrSaveBkName)
'    Set myXobjOpndBook = xRefXlsOpenClose.fncobjOpenedBook( _
'            myXbisOpnRdOnly, myXstrOpnFullName, _
'            myXbisSaveON, myXstrSaveBkName)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefXlsOpenClose()
'//xRefXlsOpenClose���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefXlsOpenClose.resetConstant
End Sub
