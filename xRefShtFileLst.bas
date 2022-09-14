Attribute VB_Name = "xRefShtFileLst"
'Includes CSrchShtCmnt
'Includes CSeriesData
'Includes PfixPickUpExistFilePathArray
'Includes PincPickUpExtensionMatchFilePathArray
'Includes PfncbisCheckFileExtension
'Includes PfixGetFileFor1DArray
'Includes PfixGetFolderFileStringInformationFor1DArray
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X�ꗗ���擾����
'Rev.003
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefShtFileLst"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXlonFileCnt As Long, myZobjFile() As Object, _
            myZstrFileName() As String, myZstrFilePath() As String, _
            myXobjFilePstdFrstCell As Object
    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
    'myZstrFileName(k) : �t�@�C����
    'myZstrFilePath(k) : �t�@�C���p�X
  Private myXstrDirPath As String, myXobjDirPstdCell As Object, _
            myXstrExtsn As String
  
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
    Call callxRefShtFileLst
    
'//�������ʕ\��
    MsgBox "�擾�p�X���F" & myXlonFileCnt
    
End Sub

'PublicP_
Public Sub callProc( _
            myXlonFileCntOUT As Long, myZobjFileOUT() As Object, _
            myZstrFileNameOUT() As String, myZstrFilePathOUT() As String, _
            myXobjFilePstdFrstCellOUT As Object, _
            myXstrDirPathOUT As String, myXobjDirPstdCellOUT As Object, _
            myXstrExtsnOUT As String, _
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
    myXlonFileCntOUT = Empty
    Erase myZobjFileOUT: Erase myZstrFileNameOUT: Erase myZstrFilePathOUT
    Set myXobjFilePstdFrstCellOUT = Nothing
    myXstrDirPathOUT = Empty: Set myXobjDirPstdCellOUT = Nothing
    myXstrExtsnOUT = Empty
    
'//�������s
    Call ctrProc
    If myXlonFileCnt <= 0 Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXlonFileCntOUT = myXlonFileCnt
    myZobjFileOUT() = myZobjFile()
    myZstrFileNameOUT() = myZstrFileName()
    myZstrFilePathOUT() = myZstrFilePath()
    Set myXobjFilePstdFrstCellOUT = myXobjFilePstdFrstCell
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
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables1
    Call setControlVariables2
    
'//S:�V�[�g��̋L�ڃf�[�^���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXlonFileCnt = Empty: Erase myZobjFile: Erase myZstrFileName: Erase myZstrFilePath
    Set myXobjFilePstdFrstCell = Nothing
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

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    
    myXlonSrchShtNo = 4
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'    Set myXobjSrchSheet = ActiveSheet

  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 3
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
    myZvarSrchCndtn(k, L + 0) = "��������t�@�C���g���q�F"
    myZvarSrchCndtn(k, L + 1) = 0
    myZvarSrchCndtn(k, L + 2) = 1
    myZvarSrchCndtn(k, L + 3) = 0
    k = k + 1   'k = 3
    myZvarSrchCndtn(k, L + 0) = "�t�@�C���ꗗ"
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
    myXstrExtsn = myZstrTrgtVal(Lc + 1)
    Set myXobjDirPstdCell = myZobjTrgtRng(Lc + 0)
    Set myXobjFilePstdFrstCell = myZobjTrgtRng(Lc + 2)
    If myXobjFilePstdFrstCell Is Nothing Then GoTo ExitPath
    
'//�t�@�C���p�X�ꗗ���擾
    myXlonBgnRow = myXobjFilePstdFrstCell.Row
    myXlonBgnCol = myXobjFilePstdFrstCell.Column
    Call instCSeriesData
    If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
    
  Dim myZstrFilePathOrg1() As String, myZstrFilePathOrg2() As String
  Dim i As Long
    i = myXlonSrsDataCnt + Lo - 1
    ReDim myZstrFilePathOrg1(i) As String
    ReDim myZstrFilePathOrg2(i) As String
    Lc = LBound(myZstrSrsData)
    For i = 1 To myXlonSrsDataCnt
        myZstrFilePathOrg1(i + Lo - 1) = myXstrDirPath & "\" & myZstrSrsData(i + Lc - 1)
        myZstrFilePathOrg2(i + Lo - 1) = myZstrSrsData(i + Lc - 1)
    Next i
    
'//�擾�����t�@�C���p�X�ꗗ���瑶�݂Ɗg���q�őI��
  Dim myXlonExistFileCnt As Long, myZstrExistFilePath() As String
  Dim myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String
    
    Call PfixPickUpExistFilePathArray( _
            myXlonExistFileCnt, myZstrExistFilePath, _
            myZstrFilePathOrg1)
    Call PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt, myZstrExtMtchFilePath, _
            myZstrExistFilePath, myXstrExtsn)
    If myXlonExtMtchFileCnt > 0 Then GoTo JumpPath
    
    Call PfixPickUpExistFilePathArray( _
            myXlonExistFileCnt, myZstrExistFilePath, _
            myZstrFilePathOrg2)
    Call PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt, myZstrExtMtchFilePath, _
            myZstrExistFilePath, myXstrExtsn)
    If myXlonExtMtchFileCnt <= 0 Then GoTo ExitPath
    
JumpPath:
'//�t�@�C���p�X�ꗗ����t�@�C���I�u�W�F�N�g�ꗗ���擾
    Call PfixGetFileFor1DArray(myXlonFileCnt, myZobjFile, myZstrExtMtchFilePath)

'//�t�@�C���ꗗ�̃t�@�C�������擾
  Dim myXlonInfoCnt As Long
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFileName, _
            myZobjFile, 1)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath

'//�t�@�C���ꗗ�̃t�@�C���p�X���擾
    Call PfixGetFolderFileStringInformationFor1DArray( _
            myXlonInfoCnt, myZstrFilePath, _
            myZobjFile, 2)
    If myXlonInfoCnt <= 0 Then GoTo ExitPath
    
    Erase myZstrFilePathOrg1: Erase myZstrFilePathOrg2
    Erase myZstrExistFilePath: Erase myZstrExtMtchFilePath
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

 '��^�o_�t�@�C���ꗗ����w��g���q�ƈ�v����t�@�C���p�X�𒊏o����
Private Sub PincPickUpExtensionMatchFilePathArray( _
            myXlonExtMtchFileCnt As Long, myZstrExtMtchFilePath() As String, _
            ByRef myXstrOrgFilePath() As String, ByVal myXstrExtsn As String)
'Includes PfncbisCheckFileExtension
'myZstrExtMtchFilePath(i) : ���o�t�@�C���p�X
'myXstrOrgFilePath(i) : ���t�@�C���p�X
    myXlonExtMtchFileCnt = Empty: Erase myZstrExtMtchFilePath
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myXstrOrgFilePath): myXstrTmp = myXstrOrgFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXvarFilePath As Variant, myXbisExtChck As Boolean, n As Long: n = Lo - 1
    For Each myXvarFilePath In myXstrOrgFilePath
      Dim myXstrFilePath As String: myXstrFilePath = myXvarFilePath
        myXbisExtChck = PfncbisCheckFileExtension(myXstrFilePath, myXstrExtsn)
        If myXbisExtChck = False Then GoTo NextPath
        n = n + 1: ReDim Preserve myZstrExtMtchFilePath(n) As String
        myZstrExtMtchFilePath(n) = myXvarFilePath
NextPath:
    Next
    myXlonExtMtchFileCnt = n - Lo + 1
    myXvarFilePath = Empty
ExitPath:
End Sub

 '��^�e_�w��t�@�C�����w��g���q�ł��邱�Ƃ��m�F����
Private Function PfncbisCheckFileExtension( _
            ByVal myXstrFilePath As String, ByVal myXstrExtsn As String) As Boolean
'myXstrExtsn = "*" : �C�ӂ̕�����̃��C���h�J�[�h
    PfncbisCheckFileExtension = False
    If myXstrFilePath = "" Then Exit Function
    If myXstrExtsn = "" Then GoTo JumpPath
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXstrOrgExt As String
    With myXobjFSO
        If .FileExists(myXstrFilePath) = False Then Exit Function
        myXstrOrgExt = .GetExtensionName(myXstrFilePath)
    End With
  Dim myXstrDesExt As String: myXstrDesExt = myXstrExtsn
    If Left(myXstrDesExt, 1) = "." Then myXstrDesExt = Mid(myXstrDesExt, 2)
    myXstrOrgExt = LCase(myXstrOrgExt)
    myXstrDesExt = LCase(myXstrDesExt)
    If myXstrOrgExt = myXstrDesExt Then GoTo JumpPath
  Dim myXlonPstn As Long: myXlonPstn = InStr(myXstrDesExt, "*")
    Select Case myXlonPstn
        Case 1
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
        Case Len(myXstrExtsn)
            If Left(myXstrOrgExt, Len(myXstrDesExt) - 1) _
                    <> Left(myXstrDesExt, Len(myXstrDesExt) - 1) Then _
                Exit Function
        Case Else
            If Right(myXstrOrgExt, Len(myXstrDesExt) - myXlonPstn) _
                    <> Right(myXstrDesExt, Len(myXstrDesExt) - myXlonPstn) Then _
                Exit Function
            If Left(myXstrOrgExt, myXlonPstn - 1) _
                    <> Left(myXstrDesExt, myXlonPstn - 1) Then _
                Exit Function
    End Select
    Set myXobjFSO = Nothing
JumpPath:
    PfncbisCheckFileExtension = True
End Function

 '��^�o_1�����z��̃t�@�C���p�X����t�@�C���I�u�W�F�N�g�ꗗ���擾����
Private Sub PfixGetFileFor1DArray( _
                myXlonFileCnt As Long, myZobjFile() As Object, _
                ByRef myZstrFilePath() As String)
'myZobjFile(i) : �t�@�C���I�u�W�F�N�g�ꗗ
'myZstrFilePath(i) : ���t�@�C���p�X�ꗗ
    myXlonFileCnt = Empty: Erase myZobjFile
  Dim myXstrTmp As String, Li As Long
    On Error GoTo ExitPath
    Li = LBound(myZstrFilePath): myXstrTmp = myZstrFilePath(Li)
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjTmp As Object, i As Long, n As Long: n = Lo - 1
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    For i = LBound(myZstrFilePath) To UBound(myZstrFilePath)
        myXstrTmp = Empty
        myXstrTmp = myZstrFilePath(i)
        With myXobjFSO
            If .FileExists(myXstrTmp) = False Then GoTo NextPath
            Set myXobjTmp = .GetFile(myXstrTmp)
        End With
        n = n + 1: ReDim Preserve myZobjFile(n) As Object
        Set myZobjFile(n) = myXobjTmp
NextPath:
    Next i
    myXlonFileCnt = n - Lo + 1
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
'    myZvarSrchCndtn(k, L + 0) = "�t�@�C���ꗗ"
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
'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X�ꗗ���擾����
Private Sub callxRefShtFileLst()
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
'  Dim myXlonFileCnt As Long, myZobjFile() As Object, _
'        myZstrFileName() As String, myZstrFilePath() As String, _
'        myXobjFilePstdFrstCell As Object, _
'        myXstrDirPath As String, myXobjDirPstdCell As Object, _
'        myXstrExtsn As String
'    'myZobjFile(k) : �t�@�C���I�u�W�F�N�g
'    'myZstrFileName(k) : �t�@�C����
'    'myZstrFilePath(k) : �t�@�C���p�X
    Call xRefShtFileLst.callProc( _
            myXlonFileCnt, myZobjFile, myZstrFileName, myZstrFilePath, _
            myXobjFilePstdFrstCell, _
            myXstrDirPath, myXobjDirPstdCell, myXstrExtsn, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn, myXbisRowDrctn)
    Debug.Print "�f�[�^: " & myXlonFileCnt
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefShtFileLst()
'//xRefShtFileLst���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefShtFileLst.resetConstant
End Sub
