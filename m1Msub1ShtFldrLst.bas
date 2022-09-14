Attribute VB_Name = "m1Msub1ShtFldrLst"
'Includes CSrchShtCmnt
'Includes CSeriesData
'Includes PfixPickUpExistFolderArray
'Includes PfixGetFolderFor1DArray
'Includes PfixGetFolderFileStringInformationFor1DArray
'Includes PfixChangeModuleConstValue

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�H���_�p�X�ꗗ���擾����
'Rev.002
  
'//���W���[��������
  Private Const meMstrMdlName As String = "m1Msub1ShtFldrLst"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonFldrCnt As Long, myZobjFldr() As Object, _
            myZstrFldrName() As String, myZstrFldrPath() As String, _
            myXobjFldrPstdFrstCell As Object
    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
    'myZstrFldrName(k) : �t�H���_��
    'myZstrFldrPath(k) : �t�H���_�p�X
  Private myXstrDirPath As String, myXobjDirPstdCell As Object
  
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
  
  Private myXbisRowDrctn As Boolean
    'myXbisRowDrctn = True  : �s�����݂̂�����
    'myXbisRowDrctn = False : ������݂̂�����
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
'    'myZstrTrgtVal(i) : �擾������
'    'myZobjTrgtRng(i) : �s��ʒu�̃Z��
  
  Private myXlonBgnRow As Long, myXlonBgnCol As Long
  
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : �擾������

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonTrgtValCnt = Empty: Erase myZstrTrgtVal: Erase myZobjTrgtRng
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
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
    Call callm1Msub1ShtFldrLst
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFldrCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFldrCntOUT As Long, myZobjFldrOUT() As Object, _
            myZstrFldrNameOUT() As String, myZstrFldrPathOUT() As String, _
            myXobjFldrPstdFrstCellOUT As Object, _
            myXstrDirPathOUT As String, myXobjDirPstdCellOUT As Object, _
            ByVal myXlonSrchShtNoIN As Long, ByVal myXobjSrchSheetIN As Object, _
            ByVal myXlonShtSrchCntIN As Long, ByRef myZvarSrchCndtnIN As Variant, _
            ByVal myXbisInStrOptnIN As Boolean, _
            ByVal myXbisRowDrctnIN As Boolean)
    
'//���͕ϐ���������
    myXlonSrchShtNo = Empty: Set myXobjSrchSheet = Nothing
    myXlonShtSrchCnt = Empty: myZvarSrchCndtn = Empty
    myXbisInStrOptn = False
    myXbisRowDrctn = False

'//���͕ϐ�����荞��
    myXlonSrchShtNo = myXlonSrchShtNoIN
    Set myXobjSrchSheet = myXobjSrchSheetIN
    myXlonShtSrchCnt = myXlonShtSrchCntIN
    myZvarSrchCndtn = myZvarSrchCndtnIN
    myXbisInStrOptn = myXbisInStrOptnIN
    myXbisRowDrctn = myXbisRowDrctnIN
    
'//�o�͕ϐ���������
    myXlonFldrCntOUT = Empty
    Erase myZobjFldrOUT: Erase myZstrFldrNameOUT: Erase myZobjFldrOUT
    Set myXobjFldrPstdFrstCellOUT = Nothing
    myXstrDirPathOUT = Empty: Set myXobjDirPstdCellOUT = Nothing
    
'//�������s
    Call ctrProc
'    If myXlonFldrCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonFldrCntOUT = myXlonFldrCnt
    myZobjFldrOUT() = myZobjFldr()
    myZstrFldrNameOUT() = myZstrFldrName()
    myZstrFldrPathOUT() = myZstrFldrPath()
    Set myXobjFldrPstdFrstCellOUT = myXobjFldrPstdFrstCell
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjDirPstdCellOUT = myXobjDirPstdCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables1
    Call setControlVariables2
    
'//S:�V�[�g��̋L�ڃf�[�^���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXlonFldrCnt = Empty: Erase myZobjFldr: Erase myZstrFldrName: Erase myZstrFldrPath
    Set myXobjFldrPstdFrstCell = Nothing
    myXstrDirPath = Empty: Set myXobjDirPstdCell = Nothing
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
Private Sub setControlVariables1()
    
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
    Set myXobjSrchSheet = ActiveSheet

  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 2
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : ����������
    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "�e�t�H���_�p�X�F"
    myZvarSrchCndtn(k, L + 1) = 0
    myZvarSrchCndtn(k, L + 2) = 1
    myZvarSrchCndtn(k, L + 3) = 0
    k = k + 1   'k = 2
    myZvarSrchCndtn(k, L + 0) = "�t�H���_�ꗗ"
    myZvarSrchCndtn(k, L + 1) = 1
    myZvarSrchCndtn(k, L + 2) = 0
    myZvarSrchCndtn(k, L + 3) = 0
    
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    myXbisRowDrctn = True
    'myXbisRowDrctn = True  : �s�����݂̂�����
    'myXbisRowDrctn = False : ������݂̂�����
    
End Sub

'SnsP_�V�[�g��̋L�ڃf�[�^���擾
Private Sub snsProc()
    myXbisExitFlag = False
    
'//�f�B���N�g���p�X���������Ď擾
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim Lc As Long: Lc = LBound(myZstrTrgtVal)
    myXstrDirPath = myZstrTrgtVal(Lc + 0)
    Set myXobjDirPstdCell = myZobjTrgtRng(Lc + 0)
    Set myXobjFldrPstdFrstCell = myZobjTrgtRng(Lc + 1)
    If myXobjFldrPstdFrstCell Is Nothing Then GoTo ExitPath
    
'//�t�H���_�p�X�ꗗ���擾
    myXlonBgnRow = myXobjFldrPstdFrstCell.Row
    myXlonBgnCol = myXobjFldrPstdFrstCell.Column
    Call instCSeriesData
    If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
    
  Dim myZstrFldrPathOrg1() As String, myZstrFldrPathOrg2() As String
  Dim i As Long
    i = myXlonSrsDataCnt + Lo - 1
    ReDim myZstrFldrPathOrg1(i) As String
    ReDim myZstrFldrPathOrg2(i) As String
    Lc = LBound(myZstrSrsData)
    For i = 1 To myXlonSrsDataCnt
        myZstrFldrPathOrg1(i + Lo - 1) = myXstrDirPath & "\" & myZstrSrsData(i + Lc - 1)
        myZstrFldrPathOrg2(i + Lo - 1) = myZstrSrsData(i + Lc - 1)
    Next i
    
'//�擾�����t�H���_�p�X�ꗗ���瑶�݂őI��
  Dim myXlonExistFldrCnt As Long, myZstrExistFldrPath() As String
    Call PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt, myZstrExistFldrPath, _
            myZstrFldrPathOrg1)
    If myXlonExistFldrCnt > 0 Then GoTo JumpPath
    
    Call PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt, myZstrExistFldrPath, _
            myZstrFldrPathOrg2)
    If myXlonExistFldrCnt <= 0 Then GoTo ExitPath
    
JumpPath:
'//�t�H���_�p�X�ꗗ����t�H���_�I�u�W�F�N�g�ꗗ���擾
    Call PfixGetFolderFor1DArray(myXlonFldrCnt, myZobjFldr, myZstrExistFldrPath)

'//�t�H���_�ꗗ�̃t�H���_�����擾
  Dim myXlonInfoCnt As Long
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFldrName, _
            myZobjFldr, 1)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath

'//�t�H���_�ꗗ�̃t�H���_�p�X���擾
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFldrPath, _
            myZobjFldr, 2)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrTrgtVal: Erase myZobjTrgtRng
    Erase myZstrFldrPathOrg1: Erase myZstrFldrPathOrg2
    Erase myZstrExistFldrPath
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

'��ClassProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
Private Sub instCSeriesData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
    With myXinsSeriesData
    '//�N���X���ϐ��ւ̓���
        Set .setSrchSheet = myXobjSrchSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
        .letRowDrctn = myXbisRowDrctn
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonSrsDataCnt = .getSrsDataCnt
        If myXlonSrsDataCnt <= 0 Then GoTo JumpPath
        k = myXlonSrsDataCnt + Lo - 1
        ReDim myZstrSrsData(k) As String
        Lc = .getOptnBase
        For k = 1 To myXlonSrsDataCnt
            myZstrSrsData(k + Lo - 1) = .getSrsDataAry(k + Lc - 1)
        Next k
    End With
JumpPath:
    Set myXinsSeriesData = Nothing
End Sub

'===============================================================================================

 '��^�o_�t�H���_�p�X�ꗗ���瑶�݂���t�H���_�p�X�𒊏o����
Private Sub PfixPickUpExistFolderArray( _
            myXlonExistFldrCnt As Long, myZstrExistFldrPath() As String, _
            ByRef myZstrOrgFldrPath() As String)
'myZstrExistFldrPath(i) : ���o�t�H���_�p�X
'myZstrOrgFldrPath(i) : ���t�H���_�p�X
    myXlonExistFldrCnt = Empty: Erase myZstrExistFldrPath
  Dim myXstrTmp As String, L As Long
    On Error GoTo ExitPath
    L = LBound(myZstrOrgFldrPath): myXstrTmp = myZstrOrgFldrPath(L)
    On Error GoTo 0
  Dim i As Long, myXstrPath As String, n As Long: n = L - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrOrgFldrPath) To UBound(myZstrOrgFldrPath)
        myXstrPath = myZstrOrgFldrPath(i)
        If myXobjFSO.FolderExists(myXstrPath) = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExistFldrPath(n) As String
        myZstrExistFldrPath(n) = myXstrPath
NextPath:
    Next i
    myXlonExistFldrCnt = n + L - 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '��^�o_1�����z��̃t�H���_�p�X�ꗗ����t�H���_�I�u�W�F�N�g�ꗗ���擾����
Private Sub PfixGetFolderFor1DArray( _
                myXlonFldrCnt As Long, myZobjFldr() As Object, _
                ByRef myZstrFldrPath() As String)
'myZobjFldr(i) : �t�H���_�I�u�W�F�N�g�ꗗ
'myZstrFldrPath(i) : ���t�H���_�p�X�ꗗ
    myXlonFldrCnt = Empty: Erase myZobjFldr
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFldrPath): myXstrTmp = myZstrFldrPath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjTmp As Object, i As Long, n As Long: n = Lo - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFldrPath) To UBound(myZstrFldrPath)
        myXstrTmp = Empty
        myXstrTmp = myZstrFldrPath(i)
        With myXobjFSO
            If .FolderExists(myXstrTmp) = False Then GoTo NextPath
            Set myXobjTmp = .GetFolder(myXstrTmp)
        End With
        n = n + 1: ReDim Preserve myZobjFldr(n) As Object
        Set myZobjFldr(n) = myXobjTmp
NextPath:
    Next i
    myXlonFldrCnt = n - Lo + 1
    Set myXobjFSO = Nothing
ExitPath:
End Sub

 '��^�o_1�����z��̃t�H���_�t�@�C���I�u�W�F�N�g�ꗗ�̕���������擾����
Private Sub PfixGetFolderFileStringInformationFor1DArray( _
                myXlonInfoCnt As Long, myZstrInfo() As String, _
                ByRef myZobjFldrFile() As Object, _
                Optional ByVal coXlonStrOptn As Long = 1)
'myZstrInfo(i) : ���o�t�H���_���
'myZobjFldrFile(i) : ���t�H���_or���t�@�C��
'coXlonStrOptn = 1  : ���O (Name)
'coXlonStrOptn = 2  : �p�X (Path)
'coXlonStrOptn = 3  : �e�t�H���_ (ParentFolder)
'coXlonStrOptn = 4  : ���� (Attributes)
'coXlonStrOptn = 5  : ��� (Type)
    myXlonInfoCnt = Empty: Erase myZstrInfo
  Dim myXobjTmp As Object, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZobjFldrFile): Set myXobjTmp = myZobjFldrFile(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXstrTmp As String, i As Long, n As Long: n = Lo - 1
    On Error GoTo NextPath
    For i = LBound(myZobjFldrFile) To UBound(myZobjFldrFile)
        myXstrTmp = Empty
        If myZobjFldrFile(i) Is Nothing Then GoTo NextPath
        Select Case coXlonStrOptn
            Case 1: myXstrTmp = myZobjFldrFile(i).Name
            Case 2: myXstrTmp = myZobjFldrFile(i).Path
            Case 3: myXstrTmp = myZobjFldrFile(i).ParentFolder
            Case 4: myXstrTmp = myZobjFldrFile(i).Attributes
            Case 5: myXstrTmp = myZobjFldrFile(i).Type
        End Select
        n = n + 1: ReDim Preserve myZstrInfo(n) As String
        myZstrInfo(n) = myXstrTmp
NextPath:
    Next i
    On Error GoTo 0
    myXlonInfoCnt = n - Lo + 1
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
'Private Sub setControlVariables1()
'    myXlonSrchShtNo = 2
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
''    Set myXobjSrchSheet = ActiveSheet
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'    myXlonShtSrchCnt = 2
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
'    myZvarSrchCndtn(k, L + 0) = "�t�H���_�ꗗ"
'    myZvarSrchCndtn(k, L + 1) = 1
'    myZvarSrchCndtn(k, L + 2) = 0
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXbisRowDrctn = True
'    'myXbisRowDrctn = True  : �s�����݂̂�����
'    'myXbisRowDrctn = False : ������݂̂�����
'End Sub
'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�H���_�p�X�ꗗ���擾����
Private Sub callm1Msub1ShtFldrLst()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'    'myZvarSrchCndtn(i, 1) : ����������
'    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
'    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
'    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'  Dim myXbisRowDrctn As Boolean
'    'myXbisRowDrctn = True  : �s�����݂̂�����
'    'myXbisRowDrctn = False : ������݂̂�����
'  Dim myXlonFldrCnt As Long, myZobjFldr() As Object, _
'        myZstrFldrName() As String, myZstrFldrPath() As String, _
'        myXobjFldrPstdFrstCell As Object, _
'        myXstrDirPath As String, myXobjDirPstdCell As Object
'    'myZobjFldr(k) : �t�H���_�I�u�W�F�N�g
'    'myZstrFldrName(k) : �t�H���_��
'    'myZstrFldrPath(k) : �t�H���_�p�X
    Call m1Msub1ShtFldrLst.callProc( _
            myXlonFldrCnt, myZobjFldr, myZstrFldrName, myZstrFldrPath, _
            myXobjFldrPstdFrstCell, _
            myXstrDirPath, myXobjDirPstdCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn, myXbisRowDrctn)
'    Debug.Print "�f�[�^: " & myXlonFldrCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInm1Msub1ShtFldrLst()
'//m1Msub1ShtFldrLst���W���[���̃��W���[���������̃��Z�b�g����
    Call m1Msub1ShtFldrLst.resetConstant
End Sub
