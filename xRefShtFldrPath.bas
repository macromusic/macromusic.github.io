Attribute VB_Name = "xRefShtFldrPath"
'Includes CSrchShtCmnt
'Includes PfncbisCheckFolderExist
'Includes PfncobjGetFolder
'Includes PfixGetFolderNameInformationByFSO
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�H���_�p�X���擾����
'Rev.003
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefShtFldrPath"
  Private Const meMlonExeNum As Long = 0
  
'//�o�̓f�[�^
  Private myXobjFldr As Object, myXstrFldrName As String, myXstrFldrPath As String, _
            myXstrDirPath As String, _
            myXobjFldrPstdFrstCell As Object
  
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
    Call callOfxRefShtFldrPath
    
'//�������ʕ\��
    MsgBox "�擾�p�X�F" & myXstrFldrPath
    
End Sub

'PublicP_
Public Sub callProc( _
            myXobjFldrOUT As Object, myXstrFldrNameOUT As String, myXstrFldrPathOUT As String, _
            myXstrDirPathOUT As String, _
            myXobjFldrPstdFrstCellOUT As Object, _
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
    Set myXobjFldrOUT = Nothing: myXstrFldrNameOUT = Empty: myXstrFldrPathOUT = Empty
    myXstrDirPathOUT = Empty
    Set myXobjFldrPstdFrstCellOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXstrFldrPath = "" Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    Set myXobjFldrOUT = myXobjFldr
    myXstrFldrNameOUT = myXstrFldrName
    myXstrFldrPathOUT = myXstrFldrPath
    myXstrDirPathOUT = myXstrDirPath
    Set myXobjFldrPstdFrstCellOUT = myXobjFldrPstdFrstCell
    
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
    myXstrFldrPath = Empty: myXstrFldrName = Empty
    myXstrDirPath = Empty: Set myXobjFldr = Nothing
    Set myXobjFldrPstdFrstCell = Nothing
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
    myZvarSrchCndtn(k, L + 0) = "�t�H���_�p�X�F"
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
    
'//�t�H���_�p�X���������Ď擾
    Call instCSrchShtCmnt
    If myXlonTrgtValCnt <= 0 Then GoTo ExitPath
    If myXlonTrgtValCnt <> myXlonShtSrchCnt Then GoTo ExitPath
    
  Dim L As Long: L = LBound(myZobjTrgtRng)
    myXstrFldrPath = myZstrTrgtVal(L)
    Set myXobjFldrPstdFrstCell = myZobjTrgtRng(L)
    If myXobjFldrPstdFrstCell Is Nothing Then GoTo ExitPath
    
'//�w��t�H���_�̑��݂��m�F
    If PfncbisCheckFolderExist(myXstrFldrPath) = False Then
        myXstrFldrPath = ""
        GoTo ExitPath
    End If
    
'//�w��t�H���_�̃I�u�W�F�N�g���擾
    Set myXobjFldr = PfncobjGetFolder(myXstrFldrPath)
    
'//�w��t�H���_�̃t�H���_�������擾
    Call PfixGetFolderNameInformationByFSO(myXstrDirPath, myXstrFldrName, myXstrFldrPath)
    
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

 '��^�e_�w��t�H���_�̑��݂��m�F����
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function

 '��^�e_�w��t�H���_�̃I�u�W�F�N�g���擾����
Private Function PfncobjGetFolder(ByVal myXstrDirPath As String) As Object
    Set PfncobjGetFolder = Nothing
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        If .FolderExists(myXstrDirPath) = False Then Exit Function
        Set PfncobjGetFolder = .GetFolder(myXstrDirPath)
    End With
    Set myXobjFSO = Nothing
End Function

 '��^�o_�w��t�H���_�̃t�H���_�������擾����(FileSystemObject�g�p)
Private Sub PfixGetFolderNameInformationByFSO( _
            myXstrPrntPath As String, myXstrDirName As String, _
            ByVal myXstrDirPath As String)
    myXstrPrntPath = Empty: myXstrDirName = Empty
    If myXstrDirPath = "" Then Exit Sub
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    With myXobjFSO
        myXstrPrntPath = .GetParentFolderName(myXstrDirPath)    '�e�t�H���_�p�X
        myXstrDirName = .GetFolder(myXstrDirPath).Name          '�t�H���_��
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
'    myZvarSrchCndtn(k, L + 0) = "�t�H���_�p�X�F"
'    myZvarSrchCndtn(k, L + 1) = 0
'    myZvarSrchCndtn(k, L + 2) = 1
'    myZvarSrchCndtn(k, L + 3) = 0
'    myXbisInStrOptn = False
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'End Sub
'��ModuleProc��_�G�N�Z���V�[�g��ɋL�ڂ��ꂽ�t�H���_�p�X���擾����
Private Sub callOfxRefShtFldrPath()
'  Dim myXlonSrchShtNo As Long, myXobjSrchSheet As Object, _
'        myXlonShtSrchCnt As Long, myZvarSrchCndtn As Variant, _
'        myXbisInStrOptn As Boolean
'    'myZvarSrchCndtn(i, 1) : ����������
'    'myZvarSrchCndtn(i, 2) : �I�t�Z�b�g�s��
'    'myZvarSrchCndtn(i, 3) : �I�t�Z�b�g��
'    'myZvarSrchCndtn(i, 4) : �V�[�g�㕶���񌟍�[=0]or�R�����g�������񌟍�[=1]
'    'myXbisInStrOptn = False : �w�蕶����ƈ�v��������Ō�������
'    'myXbisInStrOptn = True  : �w�蕶������܂ޏ����Ō�������
'  Dim myXstrFldrPath As String, myXstrFldrName As String, _
'        myXstrDirPath As String, myXobjFldr As Object, _
'        myXobjFldrPstdFrstCell As Object
    Call xRefShtFldrPath.callProc( _
            myXobjFldr, myXstrFldrName, myXstrFldrPath, myXstrDirPath, _
            myXobjFldrPstdFrstCell, _
            myXlonSrchShtNo, myXobjSrchSheet, myXlonShtSrchCnt, myZvarSrchCndtn, _
            myXbisInStrOptn)
    Debug.Print "�f�[�^: " & myXstrFldrPath
    Debug.Print "�f�[�^: " & myXstrFldrName
    Debug.Print "�f�[�^: " & myXstrDirPath
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefShtFldrPath()
'//xRefShtFldrPath���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefShtFldrPath.resetConstant
End Sub
