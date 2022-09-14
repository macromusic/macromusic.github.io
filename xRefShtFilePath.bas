Attribute VB_Name = "xRefShtFilePath"
'Includes CSrchShtCmnt
'Includes PfncbisCheckFileExist
'Includes PfncobjGetFile
'Includes PfixGetFileNameInformationByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X���擾����
'Rev.003
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefShtFilePath"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXstrFilePath As String, myXstrFileName As String, _
            myXstrDirPath As String, myXobjFile As Object, _
            myXobjFilePstdFrstCell As Object
  
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
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonTrgtValCnt As Long, myZstrTrgtVal() As String, myZobjTrgtRng() As Object
'    'myZstrTrgtVal(i) : �擾������
'    'myZobjTrgtRng(i) : �s��ʒu�̃Z��

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonTrgtValCnt = Empty: Erase myZstrTrgtVal: Erase myZobjTrgtRng
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
    Call callOfxRefShtFilePath
    
'//�������ʕ\��
    MsgBox "�擾�p�X�F" & myXstrFilePath
    
End Sub

'PublicP_
Public Sub callProc( _
            myXstrFilePathOUT As String, myXstrFileNameOUT As String, _
            myXstrDirPathOUT As String, myXobjFileOUT As Object, _
            myXobjFilePstdFrstCellOUT As Object, _
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
    myXstrFilePathOUT = Empty: myXstrFileNameOUT = Empty
    myXstrDirPathOUT = Empty: Set myXobjFileOUT = Nothing
    Set myXobjFilePstdFrstCellOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXstrFilePath = "" Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXstrFilePathOUT = myXstrFilePath
    Set myXobjFileOUT = myXobjFile
    myXstrDirPathOUT = myXstrDirPath
    myXstrFileNameOUT = myXstrFileName
    Set myXobjFilePstdFrstCellOUT = myXobjFilePstdFrstCell
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
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
    myXstrFilePath = Empty: myXstrFileName = Empty
    myXstrDirPath = Empty: Set myXobjFile = Nothing
    Set myXobjFilePstdFrstCell = Nothing
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
    
    myXlonSrchShtNo = 4
    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
'    Set myXobjSrchSheet = ActiveSheet

  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    myXlonShtSrchCnt = 1
    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
    'myZvarSrchCndtn(i, 1) : ����������
    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
  Dim k As Long: k = L - 1
    k = k + 1   'k = 1
    myZvarSrchCndtn(k, L + 0) = "�t�@�C���p�X�F"
    myZvarSrchCndtn(k, L + 1) = 0
    myZvarSrchCndtn(k, L + 2) = 1
    myZvarSrchCndtn(k, L + 3) = 0
    
    myXbisInStrOptn = False
    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
    
End Sub

'SnsP_�V�[�g��̋L�ڃf�[�^���擾
Private Sub snsProc()
    myXbisExitFlag = False
    
'//�t�@�C���p�X���������Ď擾
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
  Dim L As Long: L = LBound(myZobjTrgtRng)
    myXstrFilePath = myZstrTrgtVal(L)
    Set myXobjFilePstdFrstCell = myZobjTrgtRng(L)
    If myXobjFilePstdFrstCell Is Nothing Then GoTo ExitPath
    
'//�w��t�@�C���̑��݂��m�F
    If PfncbisCheckFileExist(myXstrFilePath) = False Then
        myXstrFilePath = ""
        GoTo ExitPath
    End If
    
'//�w��t�@�C���̃I�u�W�F�N�g���擾
    Set myXobjFile = PfncobjGetFile(myXstrFilePath)
    
'//�w��t�@�C���̃t�@�C���������擾
  Dim myXstrBaseName As String, myXstrExtsn As String
    Call PfixGetFileNameInformationByFSO( _
            myXstrDirPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXstrFilePath)
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�擾�f�[�^���e���`�F�b�N
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

'===============================================================================================
 
 '��^�e_�w��t�@�C���̑��݂��m�F����
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '��^�e_�w��t�@�C���̃I�u�W�F�N�g���擾����
Private Function PfncobjGetFile(ByVal myXstrFilePath As String) As Object
    Set PfncobjGetFile = Nothing
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FileExists(myXstrFilePath) = False Then Exit Function
        Set PfncobjGetFile = .GetFile(myXstrFilePath)
    End With
    Set myXobjFSO = Nothing
End Function

 '��^�o_�w��t�@�C���̃t�@�C���������擾����(FileSystemObject�g�p)
Private Sub PfixGetFileNameInformationByFSO( _
            myXstrPrntPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String, _
            ByVal myXstrFilePath As String)
    myXstrPrntPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    If myXstrFilePath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExist As Boolean
    With myXobjFSO
        myXbisFileExist = .FileExists(myXstrFilePath)
        If myXbisFileExist = False Then Exit Sub
        myXstrPrntPath = .GetParentFolderName(myXstrFilePath)   '�e�t�H���_�p�X
        myXstrFileName = .getFileName(myXstrFilePath)           '�t�@�C����
        myXstrBaseName = .GetBaseName(myXstrFilePath)           '�t�@�C���x�[�X��
        myXstrExtsn = .GetExtensionName(myXstrFilePath)         '�t�@�C���g���q
    End With
    Set myXobjFSO = Nothing
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
'    myXlonSrchShtNo = 3
'    Set myXobjSrchSheet = ThisWorkbook.Worksheets(myXlonSrchShtNo)
''    Set myXobjSrchSheet = ActiveSheet
'  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'    myXlonShtSrchCnt = 1
'    ReDim myZvarSrchCndtn(myXlonShtSrchCnt + L - 1, L + 3) As Variant
'    'myZvarSrchCndtn(i, 1) : ����������
'    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
'    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
'    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
'  Dim k As Long: k = L - 1
'    k = k + 1   'k = 1
'    myZvarSrchCndtn(k, L + 0) = "�t�@�C���p�X�F"
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'End Sub
'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�@�C���p�X���擾����
Private Sub callOfxRefShtFilePath()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'    'myZvarSrchCndtn(i, 1) : ����������
'    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
'    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
'    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'  Dim myXstrFilePath As String, myXstrFileName As String, _
'        myXstrDirPath As String, myXobjFile As Object, _
'        myXobjFilePstdFrstCell As Object
    Call xRefShtFilePath.callProc( _
            myXstrFilePath, myXstrFileName, myXstrDirPath, myXobjFile, _
            myXobjFilePstdFrstCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn)
    Debug.Print "�f�[�^: " & myXstrFilePath
    Debug.Print "�f�[�^: " & myXstrFileName
    Debug.Print "�f�[�^: " & myXstrDirPath
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefShtFilePath()
'//xRefShtFilePath���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefShtFilePath.resetConstant
End Sub
