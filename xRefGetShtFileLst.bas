Attribute VB_Name = "xRefGetShtFileLst"
'Includes CSlctShtSrsData
'Includes CSlctShtDscrtData
'Includes PfixPickUpExistFilePathArray
'Includes PfixGetFileNameArrayByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�f�[�^��I�����ăp�X�ꗗ���擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefGetShtFileLst"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
  
'//���̓f�[�^
  Private myXbisByDscrt As Boolean
  Private myXlonRngOptn As Long
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String
    'myXbisByDscrt = False : �V�[�g��̘A���͈͂��w�肵�Ď擾����
    'myXbisByDscrt = True  : �V�[�g��̕s�A���͈͂��w�肵�Ď擾����
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonDataRowCnt As Long, myXlonDataColCnt As Long, myXlonDataCnt As Long, _
            myZstrShtData() As String, myZvarShtData() As Variant
  Private myZstrFilePathOrg() As String

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonDataRowCnt = Empty: myXlonDataColCnt = Empty: myXlonDataCnt = Empty
    Erase myZstrShtData: Erase myZvarShtData
    Erase myZstrFilePathOrg
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
    Call callxRefGetShtFileLst
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFileCntOUT As Long, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String)
    
'//�o�͕ϐ���������
    myXlonFileCntOUT = Empty: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    
'//�������s
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonFileCntOUT = myXlonFileCnt
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//S:�V�[�g��̋L�ڃf�[�^���擾
    Select Case myXbisByDscrt
        Case True: Call snsProc2
        Case Else: Call snsProc1
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:�擾�f�[�^���e���`�F�b�N
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty: Erase myZstrFileName: Erase myZstrFilePath
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

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
    
    myXbisByDscrt = True
    'myXbisByDscrt = False : �V�[�g��̘A���͈͂��w�肵�Ď擾����
    'myXbisByDscrt = True  : �V�[�g��̕s�A���͈͂��w�肵�Ď擾����
    
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
    
    myXstrInptBxPrmpt = "�����������t�@�C���p�X��I�����ĉ������B"
    myXstrInptBxTtl = "�t�@�C���p�X�̑I��"
    
End Sub

'SnsP_�V�[�g��̋L�ڃf�[�^���擾
Private Sub snsProc1()
    myXbisExitFlag = False
    
'//�V�[�g��̘A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾
    Call instCSlctShtSrsData
    If myXlonDataRowCnt <= 0 Or myXlonDataColCnt <= 0 Then GoTo ExitPath
    
  Dim i As Long, j As Long, k As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonDataCnt = myXlonDataRowCnt * myXlonDataColCnt
    k = myXlonDataCnt + L - 1
    ReDim myZstrFilePathOrg(k) As String
    k = L - 1
    For j = LBound(myZstrShtData, 2) To UBound(myZstrShtData, 2)
        For i = LBound(myZstrShtData, 1) To UBound(myZstrShtData, 1)
            k = k + 1
            myZstrFilePathOrg(k) = myZstrShtData(i, j)
        Next i
    Next j
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_�V�[�g��̋L�ڃf�[�^���擾
Private Sub snsProc2()
    myXbisExitFlag = False
  
'//�V�[�g��̕s�A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾
    Call instCSlctShtDscrtData
    If myXlonDataCnt <= 0 Then GoTo ExitPath
    
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim k As Long
    k = UBound(myZvarShtData, 1)
    ReDim myZstrFilePathOrg(k) As String
    For k = LBound(myZvarShtData, 1) To UBound(myZvarShtData, 1)
        myZstrFilePathOrg(k) = CStr(myZvarShtData(k, L + 2))
    Next k
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�擾�f�[�^���e���`�F�b�N
Private Sub prsProc()
    myXbisExitFlag = False
    
'//�t�@�C���p�X�ꗗ���瑶�݂���t�@�C���p�X�𒊏o
    Call PfixPickUpExistFilePathArray(myXlonFileCnt, myZstrFilePath, myZstrFilePathOrg)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
'//�w��t�@�C���p�X�ꗗ�̃t�@�C�����ꗗ���擾
    Call PfixGetFileNameArrayByFSO(myXlonFileCnt, myZstrFileName, myZstrFilePath)
    If myXlonFileCnt <= 0 Then GoTo ExitPath
    
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

'��ClassProc��_�V�[�g��̘A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
Private Sub instCSlctShtSrsData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSlctShtSrsData As CSlctShtSrsData: Set myXinsSlctShtSrsData = New CSlctShtSrsData
    With myXinsSlctShtSrsData
    '//�N���X���ϐ��ւ̓���
        .letRngOptn = myXlonRngOptn
        .letByVrnt = False
        .letGetCmnt = False
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonDataRowCnt = .getDataRowCnt
        myXlonDataColCnt = .getDataColCnt
        If myXlonDataRowCnt <= 0 Or myXlonDataColCnt <= 0 Then GoTo JumpPath
        i = myXlonDataRowCnt + Lo - 1: j = myXlonDataColCnt + Lo - 1
        ReDim myZstrShtData(i, j) As String
        Lc = .getOptnBase
        For j = 1 To myXlonDataColCnt
            For i = 1 To myXlonDataRowCnt
                myZstrShtData(i + Lo - 1, j + Lo - 1) _
                    = .getStrShtDataAry(i + Lc - 1, j + Lc - 1)
            Next i
        Next j
    End With
JumpPath:
    Set myXinsSlctShtSrsData = Nothing
End Sub

'��ClassProc��_�V�[�g��̕s�A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
Private Sub instCSlctShtDscrtData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsSlctShtDscrtData As CSlctShtDscrtData
    Set myXinsSlctShtDscrtData = New CSlctShtDscrtData
    With myXinsSlctShtDscrtData
    '//�N���X���ϐ��ւ̓���
        .letByVrnt = False
        .letGetCmnt = False
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonDataCnt = .getDataCnt
        If myXlonDataCnt <= 0 Then GoTo JumpPath
        i = myXlonDataCnt + Lo - 1
        ReDim myZvarShtData(i, Lo + 3) As Variant
        Lc = .getOptnBase
        For i = 1 To myXlonDataCnt
            myZvarShtData(i + Lo - 1, Lo + 2) = .getShtDataAry(i + Lc - 1, Lc + 2)
        Next i
    End With
JumpPath:
    Set myXinsSlctShtDscrtData = Nothing
End Sub

'===============================================================================================

 '��^�o_�t�@�C���p�X�ꗗ���瑶�݂���t�@�C���p�X�𒊏o����
Private Sub PfixPickUpExistFilePathArray( _
            myXlonExistFileCnt As Long, myZstrExistFilePath() As String, _
            ByRef myZstrOrgFilePath() As String)
'myZstrExistFilePath(i) : ���o�t�@�C���p�X
'myZstrOrgFilePath(i) : ���t�@�C���p�X
    myXlonExistFileCnt = Empty: Erase myZstrExistFilePath
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrOrgFilePath): myXstrTmp = myZstrOrgFilePath(Li)
    On Error GoTo 0
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarPath As Variant, myXbisExistChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarPath In myZstrOrgFilePath
        myXbisExistChck = myXobjFSO.FileExists(myXvarPath)
        If myXbisExistChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExistFilePath(n) As String
        myZstrExistFilePath(n) = CStr(myXvarPath)
NextPath:
    Next myXvarPath
    myXlonExistFileCnt = n - Lo + 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '��^�o_�w��t�@�C���p�X�ꗗ�̃t�@�C�����ꗗ���擾����(FileSystemObject�g�p)
Private Sub PfixGetFileNameArrayByFSO( _
            myXlonFileCnt As Long, myZstrFileName() As String, _
            ByRef myZstrFilePath() As String)
'myZstrFileName(i) : �t�@�C����
'myZstrFilePath(i) : �t�@�C���p�X
    myXlonFileCnt = Empty: Erase myZstrFileName
  Dim myXstrTmp As String, Li As Long, Ui As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFilePath): Ui = UBound(myZstrFilePath)
    myXstrTmp = myZstrFilePath(Li)
    On Error GoTo 0
    myXlonFileCnt = Ui - Li + 1
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, myXbisFileExist As Boolean
    i = myXlonFileCnt + Lo - 1: ReDim myZstrFileName(i) As String
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath) To UBound(myZstrFilePath)
        myXbisFileExist = myXobjFSO.FileExists(myZstrFilePath(i))
        If myXbisFileExist = True Then _
            myZstrFileName(i) = myXobjFSO.getFileName(myZstrFilePath(i))
    Next i
    Set myXobjFSO = Nothing
ExitPath:
End Sub

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
'    myXbisByDscrt = False
'    'myXbisByDscrt = False : �V�[�g��̘A���͈͂��w�肵�Ď擾����
'    'myXbisByDscrt = True  : �V�[�g��̕s�A���͈͂��w�肵�Ď擾����
'    myXlonRngOptn = 0
'    myXlonRngOptn = 0
'    'myXlonRngOptn = 0  : �I��͈�
'    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
'    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
'    'myXlonRngOptn = 3  : �S�f�[�^�͈�
'    myXstrInptBxPrmpt = "�����������t�@�C���p�X��I�����ĉ������B"
'    myXstrInptBxTtl = "�t�@�C���p�X�̑I��"
'End Sub
'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�f�[�^��I�����ăp�X�ꗗ���擾����
Private Sub callxRefGetShtFileLst()
'  Dim myXlonFileCnt As Long, myZstrFileName() As String, myZstrFilePath() As String
'    'myZstrFileName(k) : �t�@�C����
'    'myZstrFilePath(k) : �t�@�C���p�X
    Call xRefGetShtFileLst.callProc( _
            myXlonFileCnt, myZstrFileName, myZstrFilePath)
    Call variablesOfxRefGetShtFileLst(myXlonFileCnt, myZstrFilePath) 'Debug.Print
End Sub
Private Sub variablesOfxRefGetShtFileLst( _
            myXlonDataCnt As Long, myXvarField As Variant)
'//xRefGetShtFileLst������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt = 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myXvarField) To UBound(myXvarField)
        Debug.Print "�f�[�^" & k & ": " & myXvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefGetShtFileLst()
'//xRefGetShtFileLst���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefGetShtFileLst.resetConstant
End Sub
