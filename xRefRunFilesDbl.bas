Attribute VB_Name = "xRefRunFilesDbl"
'Includes PfncstrFileNameByFSO
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckArrayDimensionLength
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_2��ނ̕����t�@�C���ɑ΂��ĘA�����������{����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunFilesDbl"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonExeOrgFileCnt As Long, _
            myZstrExeOrgFileName() As String, myZstrExeOrgFilePath() As String
  Private myXlonExeSubFileCnt As Long, _
            myZstrExeSubFileName() As String, myZstrExeSubFilePath() As String
  
'//���͐���M��
  
'//���̓f�[�^
  Private myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
  Private myXlonSubFileCnt As Long, myZstrSubFilePath() As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
'  Private myZstrOrgFilePathINT() As String, myZstrSubFilePathINT() As String
  Private myXlonOrgFileNo As Long, _
            myXstrOrgFileName As String, myXstrOrgFilePath As String
  Private myXlonSubFileNo As Long, _
            myXstrSubFileName As String, myXstrSubFilePath As String

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
'    Erase myZstrOrgFilePathINT: Erase myZstrSubFilePathINT
'    myXlonOrgFileNo = Empty: myXstrOrgFileName = Empty: myXstrOrgFilePath = Empty
'    myXlonSubFileNo = Empty: myXstrSubFileName = Empty: myXstrSubFilePath = Empty
End Sub

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariablesSub()
'    myXlonSubFileNo = Empty: myXstrSubFileName = Empty: myXstrSubFilePath = Empty
End Sub

'-----------------------------------------------------------------------------------------------

'PublicP_���W���[���������̃��Z�b�g
Public Sub resetConstant()
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum": myZvarM(1, 2) = 0
'    myZvarM(2, 1) = "meMvarField": myZvarM(2, 2) = Chr(34) & Chr(34)
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
End Sub

'PublicP_
Public Sub exeProc()

'//�v���O�����\��
    '����: -
    '����: -
    '�o��: -

'    myXlonOrgFileCnt = 1
'    ReDim myZstrOrgFilePath(myXlonOrgFileCnt) As String
'    myXlonSubFileCnt = 1
'    ReDim myZstrSubFilePath(myXlonSubFileCnt) As String
    
'//�������s
    Call callxRefRunFilesDbl
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonExeOrgFileCntOUT As Long, myZstrExeOrgFileNameOUT() As String, _
            myZstrExeOrgFilePathOUT() As String, _
            myXlonExeSubFileCntOUT As Long, myZstrExeSubFileNameOUT() As String, _
            myZstrExeSubFilePathOUT() As String, _
            ByVal myXlonOrgFileCntIN As Long, ByRef myZstrOrgFilePathIN() As String, _
            ByVal myXlonSubFileCntIN As Long, ByRef myZstrSubFilePathIN() As String)
    
'//���͕ϐ���������
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
    myXlonSubFileCnt = Empty: Erase myZstrSubFilePath

'//���͕ϐ�����荞��
    If myXlonOrgFileCntIN <= 0 Then Exit Sub
    myXlonOrgFileCnt = myXlonOrgFileCntIN
    myZstrOrgFilePath() = myZstrOrgFilePathIN()
    
    If myXlonSubFileCntIN <= 0 Then Exit Sub
    myXlonSubFileCnt = myXlonSubFileCntIN
    myZstrSubFilePath() = myZstrSubFilePathIN()
    
'//�������s
    Call ctrProc(myXlonOrgFileCntIN, myZstrOrgFilePathIN, _
                    myXlonSubFileCntIN, myZstrSubFilePathIN)
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    myXlonExeOrgFileCntOUT = Empty
    Erase myZstrExeOrgFileNameOUT: Erase myZstrExeOrgFilePathOUT
    myXlonExeSubFileCntOUT = Empty
    Erase myZstrExeSubFileNameOUT: Erase myZstrExeSubFilePathOUT
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonOrgExeFileCnt <= 0 Then GoTo JumpPath
    myXlonExeOrgFileCntOUT = myXlonOrgExeFileCnt
    myZstrExeOrgFileNameOUT() = myZstrExeOrgFileName()
    myZstrExeOrgFilePathOUT() = myZstrExeOrgFilePath()
    
JumpPath:
    If myXlonExeSubFileCnt <= 0 Then Exit Sub
    myXlonExeSubFileCntOUT = myXlonExeSubFileCnt
    myZstrExeSubFileNameOUT() = myZstrExeSubFileName()
    myZstrExeSubFilePathOUT() = myZstrExeSubFilePath()
    
ExitPath:
End Sub

'PublicF_
Public Function fncbisCmpltFlag( _
                ByVal myXlonOrgFileCntIN As Long, _
                ByRef myZstrOrgFilePathIN() As String, _
                ByVal myXlonSubFileCntIN As Long, _
                ByRef myZstrSubFilePathIN() As String) As Boolean
    fncbisCmpltFlag = Empty
    
'//���͕ϐ���������
    myXlonOrgFileCnt = Empty: Erase myZstrOrgFilePath
    myXlonSubFileCnt = Empty: Erase myZstrSubFilePath

'//���͕ϐ�����荞��
    myXlonOrgFileCnt = myXlonOrgFileCntIN
    If myXlonOrgFileCnt <= 0 Then Exit Sub
    
    myXlonSubFileCnt = myXlonSubFileCntIN
    If myXlonSubFileCnt <= 0 Then Exit Sub
    
    myZstrOrgFilePath() = myZstrOrgFilePathIN()
    myZstrSubFilePath() = myZstrSubFilePathIN()
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Function
    
    fncbisCmpltFlag = myXbisCmpltFlag
    
ExitPath:
End Function

'-----------------------------------------------------------------------------------------------
'Control  : ���[�U������͂��󂯎���Ă��̓��e�ɉ�����Sense�AProcess�ARun�𐧌䂷��
'Sense    : Process�Ŏ��s���鉉�Z�����p�̃f�[�^���擾����
'Process  : Sense�Ŏ擾�����f�[�^���g�p���ĉ��Z����������
'Run      : Process�̏������ʂ��󂯂ĉ�ʕ\���Ȃǂ̏o�͏���������
'Remember : �L�^�������e��K�v�ɉ����Ď��o���ď����Ɋ��p����
'Record   : Sense�AProcess�ARun�Ŏ��s�����v���O�����ŏd�v�ȓ��e���L�^����
'-----------------------------------------------------------------------------------------------

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"    'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//S:Loop�O�̏��擾����
    Call snsProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"    'PassFlag
    
'//P:Loop�O�̏����H����
    Call prsProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"    'PassFlag
    
'//C:�t�@�C�����X�g1���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim k As Long
    For k = LBound(myZstrOrgFilePath) To UBound(myZstrOrgFilePath)
        myXstrFilePath = Empty: myXstrFileName = Empty
        myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
 
    '//C:�t�@�C�����X�g2���������s
        Call ctrSubProc
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "4-" & k   'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeOrgFileName(n) As String
        ReDim Preserve myZstrExeOrgFilePath(n) As String
        myZstrExeOrgFileName(n) = myXstrFileName
        myZstrExeOrgFilePath(n) = myXstrFilePath
NextPath:
    Next k
    myXlonExeOrgFileCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"    'PassFlag
    
'//P:Loop��̉��H����
    Call prsProcAfterLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "6"    'PassFlag

'//Run:�t�@�C�i���C�Y����
    Call runFinalize
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "7"    'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'CtrlP_�t�@�C�����X�g2���������s
Private Sub ctrSubProc()
    Call initializeModuleVariablesSub
    
'//S:Loop�O�̏��擾����
    Call snsSubProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s1"    'PassFlag
    
'//P:Loop�O�̏����H����
    Call prsSubProcBeforeLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s2"    'PassFlag
    
'//C:�t�@�C�����X�g���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim k As Long
    For k = LBound(myXstrSubFilePath) To UBound(myXstrSubFilePath)
        myXstrFilePath = Empty: myXstrFileName = Empty
        myXlonFileNo = k
        myXstrFilePath = myXstrSubFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
 
    '//S:�e�t�@�C���̃f�[�^�擾����
        Call snsSubProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "s3-" & k   'PassFlag
 
    '//P:�e�t�@�C���̃f�[�^���H����
        Call prsSubProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "s4-" & k   'PassFlag
            
    '//Run:�e�t�@�C���̃f�[�^�o�͏���
        Call runSubProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "s5-" & k   'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeSubFileName(n) As String
        ReDim Preserve myZstrExeSubFilePath(n) As String
        myZstrExeSubFileName(n) = myXstrSubFileName
        myZstrExeSubFilePath(n) = myXstrSubFilePath
NextPath:
    Next k
    myXlonExeSubFileCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "s6"    'PassFlag
    
'//P:Loop��̉��H����
    Call prsSubProcAfterLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s7"    'PassFlag
    
ExitPath:
    Call initializeModuleVariablesSub
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
'    myXbisCmpltFlag = False
'    myXlonExeOrgFileCnt = Empty
'    Erase myZstrExeOrgFileName: Erase myZstrExeOrgFilePath
'    myXlonExeSubFileCnt = Empty
'    Erase myZstrExeSubFileName: Erase myZstrExeSubFilePath
End Sub

'RemP_�ۑ������ϐ������o��
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
'    If myXvarFieldIN = Empty Then myXvarFieldIN = meMvarField
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_���͕ϐ����e���m�F����
Private Sub checkInputVariables()
    myXbisExitFlag = False
    
'  Dim Li As Long, myXstrTmp As String
'    On Error GoTo ExitPath
'    Li = LBound(myZstrOrgFilePath): myXstrTmp = myZstrOrgFilePath(Li)
'    Li = LBound(myZstrSubFilePath): myXstrTmp = myZstrSubFilePath(Li)
'    On Error GoTo 0
    
'//���͔z��ϐ�������z��ϐ��ɓ���ւ���
'  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
'  Dim Lia As Long, Uia As Long, Lib As Long, Uib As Long, i As Long
'    On Error GoTo ExitPath
'    Lia = LBound(myZstrOrgFilePath): Uia = UBound(myZstrOrgFilePath)
'    Lib = LBound(myZstrSubFilePath): Uib = UBound(myZstrSubFilePath)
'    If Uia <> Uib Then GoTo ExitPath
'    i = Ui + Lo - Li
'    ReDim myZstrOrgFilePathINT(i) As String
'    ReDim myZstrSubFilePathINT(i) As String
'    For i = Li To Ui
'        myZstrOrgFilePathINT(i + Lo - Li, j + Lo - Li) = myZstrOrgFilePath(i, j)
'        myZstrSubFilePathINT(i + Lo - Li, j + Lo - Li) = myZstrSubFilePath(i, j)
'    Next i
'    On Error GoTo 0
    
'//���͔z��ϐ��̓��e���m�F
'    If PfncbisCheckArrayDimension(myZstrOrgFilePath, 1) = False Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrSubFilePath, 1) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'CtrlP_
Private Sub ctrRunFiles()

'//C:�t�@�C�����X�g���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim myXvarTmpPath As Variant, k As Long: k = Li - 1
    For Each myXvarTmpPath In myZstrOrgFilePath
        myXstrFilePath = Empty: myXstrFileName = Empty
        k = k + 1: myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
        'XarbProgCode
        n = n + 1
        ReDim Preserve myZstrExeOrgFileName(n) As String
        ReDim Preserve myZstrExeOrgFilePath(n) As String
        myZstrExeOrgFileName(n) = myXstrFileName
        myZstrExeOrgFilePath(n) = myXstrFilePath
NextPath:
    Next myXvarTmpPath
    myXlonExeOrgFileCnt = n - Lo + 1
    
'//C:�t�@�C�����X�g���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgFilePath)
  Dim k As Long
    For k = LBound(myZstrOrgFilePath) To UBound(myZstrOrgFilePath)
        myXstrFilePath = Empty: myXstrFileName = Empty
        myXlonFileNo = k
        myXstrFilePath = myZstrOrgFilePath(k)
        myXstrFileName = PfncstrFileNameByFSO(myXstrFilePath)
        If myXstrFileName = "" Then GoTo NextPath
        'XarbProgCode
        n = n + 1
        ReDim Preserve myZstrExeFileName(n) As String
        ReDim Preserve myZstrExeFilePath(n) As String
        myZstrExeFileName(n) = myXstrFileName
        myZstrExeFilePath(n) = myXstrFilePath
NextPath:
    Next k
    myXlonExeFileCnt = n - Lo + 1
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
End Sub

'SnsP_Loop�O�̏��擾����
Private Sub snsProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop�O�̏����H����
Private Sub prsProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_Loop�O�̏��擾����
Private Sub snsSubProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s1-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop�O�̏����H����
Private Sub prsSubProcBeforeLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s2-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SnsP_�e�t�@�C���̃f�[�^�擾����
Private Sub snsSubProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s3-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�e�t�@�C���̃f�[�^���H����
Private Sub prsSubProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s4-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�e�t�@�C���̃f�[�^�o�͏���
Private Sub runSubProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s5-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop��̉��H����
Private Sub prsSubProcAfterLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "s7-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop��̉��H����
Private Sub prsProcAfterLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "6-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�t�@�C�i���C�Y����
Private Sub runFinalize()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "7-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_�o�͕ϐ����e���m�F����
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
'    If myXlonExeOrgFileCnt <= 0 Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeOrgFileName, 1) = False Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeOrgFilePath, 1) = False Then GoTo ExitPath

'    If myXlonExeSubFileCnt <= 0 Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeSubFileName, 1) = False Then GoTo ExitPath
'    If PfncbisCheckArrayDimension(myZstrExeSubFilePath, 1) = False Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RecP_�g�p�����ϐ���ۑ�����
Private Sub recProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
  Dim myZvarM(1, 2) As Variant
    myZvarM(1, 1) = "meMlonExeNum"
    myZvarM(1, 2) = meMlonExeNum + 1
'    myZvarM(1, 1) = "meMvarField"
'    myZvarM(1, 2) = myXvarField
    
  Dim myXstrMdlName As String: myXstrMdlName = meMstrMdlName
    Call PfixChangeModuleConstValue(myXbisExitFlag, myXstrMdlName, myZvarM)
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'===============================================================================================

 '��^�e_�w��t�@�C���̃t�@�C�������擾����(FileSystemObject�g�p)
Private Function PfncstrFileNameByFSO(ByVal myXstrFilePath As String) As String
    PfncstrFileNameByFSO = Empty
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
  Dim myXbisFileExist As Boolean: myXbisFileExist = myXobjFSO.FileExists(myXstrFilePath)
    If myXbisFileExist = False Then Exit Function
    PfncstrFileNameByFSO = myXobjFSO.GetFileName(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '��^�e_�z��ϐ��̎��������w�莟���ƈ�v���邩���`�F�b�N����
Private Function PfncbisCheckArrayDimension( _
            ByRef myZvarOrgData As Variant, ByVal myXlonDmnsn As Long) As Boolean
    PfncbisCheckArrayDimension = False
    If IsArray(myZvarOrgData) = False Then Exit Function
    If myXlonDmnsn <= 0 Then Exit Function
  Dim myXlonTmp As Long, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXlonTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    If k - 1 <> myXlonDmnsn Then Exit Function
    PfncbisCheckArrayDimension = True
End Function

 '��^�e_�z��ϐ��̎������Ɣz�񒷂��w��l�𖞑����邩���`�F�b�N����
Private Function PfncbisCheckArrayDimensionLength( _
            ByRef myZvarOrgData As Variant, ByVal myXlonChckAryDmnsn As Long, _
            ByRef myXlonChckAryLen() As Long) As Boolean
'myXlonChckAryDmnsn  : �z��̎������̎w��l
'myXlonChckAryLen(i) : i�����ڂ̔z�񒷂̎w��l
'myXlonChckAryLen(i) = 0 : �z�񒷂̃`�F�b�N�����{���Ȃ�
    PfncbisCheckArrayDimensionLength = False
    If myXlonChckAryDmnsn <= 0 Then Exit Function
  Dim Li As Long, Ui As Long, myXlonChckAryLenCnt As Long
    On Error Resume Next
    Li = LBound(myXlonChckAryLen): Ui = UBound(myXlonChckAryLen)
    If Err.Number = 9 Then Exit Function
    On Error GoTo 0
    myXlonChckAryLenCnt = Ui - Li + 1
    If myXlonChckAryLenCnt <= 0 Then Exit Function
  Dim i As Long
    For i = LBound(myXlonChckAryLen) To UBound(myXlonChckAryLen)
        If myXlonChckAryLen(i) <= 0 Then Exit Function
    Next i
'//�z��ł��邱�Ƃ��m�F
    If IsArray(myZvarOrgData) = False Then Exit Function
'//�z�񂪋�łȂ����Ƃ��m�F
  Dim myXlonTmp As Long
    On Error Resume Next
    myXlonTmp = UBound(myZvarOrgData) - LBound(myZvarOrgData) + 1
    If Err.Number = 9 Then Exit Function
    On Error GoTo 0
    If myXlonTmp <= 0 Then Exit Function
'//�z��̎��������擾
  Dim myXlonAryDmnsn As Long, myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    myXlonAryDmnsn = k - 1
    If myXlonAryDmnsn <> myXlonChckAryDmnsn Then Exit Function
    If myXlonAryDmnsn <> myXlonChckAryLenCnt Then Exit Function
'//�z��̍ŏ��Y���ƍő�Y�����擾
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
    i = myXlonAryDmnsn + L - 1
  Dim myZlonAryLBnd() As Long: ReDim myZlonAryLBnd(i) As Long
  Dim myZlonAryUBnd() As Long: ReDim myZlonAryUBnd(i) As Long
    k = 0
    For i = LBound(myZlonAryLBnd) To UBound(myZlonAryLBnd)
        k = k + 1
        myZlonAryLBnd(i) = LBound(myZvarOrgData, k)
        myZlonAryUBnd(i) = UBound(myZvarOrgData, k)
    Next i
'//�z�񒷂��擾
    i = myXlonAryDmnsn + L - 1
  Dim myZlonAryLen() As Long: ReDim myZlonAryLen(i) As Long
    For i = LBound(myZlonAryLen) To UBound(myZlonAryLen)
        myZlonAryLen(i) = myZlonAryUBnd(i) - myZlonAryLBnd(i) + 1
    Next i
'//�������Ɣz�񒷂��`�F�b�N
    For i = LBound(myZlonAryLen) To UBound(myZlonAryLen)
        If myXlonChckAryLen(i + Li - L) <> 0 Then _
            If myZlonAryLen(i) <> myXlonChckAryLen(i + Li - L) Then Exit Function
    Next i
    PfncbisCheckArrayDimensionLength = True
    Erase myZlonAryLBnd: Erase myZlonAryUBnd: Erase myZlonAryLen
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

'��ModuleProc��_2��ނ̕����t�@�C���ɑ΂��ĘA�����������{����
Private Sub callxRefRunFilesDbl()
  Dim myXlonOrgFileCnt As Long, myZstrOrgFilePath() As String
  Dim myXlonSubFileCnt As Long, myZstrSubFilePath() As String
    'myZstrOrgFilePath(i) : ���t�@�C���p�X1
    'myZstrSubFilePath(i) : ���t�@�C���p�X2
    myXlonOrgFileCnt = XarbLong
    ReDim myZstrOrgFilePath(1) As String
    myZstrOrgFilePath(1) = XarbString
    myXlonSubFileCnt = XarbLong
    ReDim myZstrSubFilePath(1) As String
    myZstrSubFilePath(1) = XarbString
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonExeOrgFileCnt As Long, _
'        myZstrExeOrgFileName() As String, myZstrOrgExeFilePath() As String
'    'myZstrExeOrgFileName(i) : ���s�t�@�C����1
'    'myZstrOrgExeFilePath(i) : ���s�t�@�C���p�X1
'  Dim myXlonExeSubFileCnt As Long, _
'        myZstrExeSubFileName() As String, myZstrSubExeFilePath() As String
'    'myZstrExeSubFileName(i) : ���s�t�@�C����2
'    'myZstrSubExeFilePath(i) : ���s�t�@�C���p�X2
    Call xRefRunFilesDbl.callProc( _
            myXbisCmpltFlag, _
            myXlonExeOrgFileCnt, myZstrExeOrgFileName, myZstrExeOrgFilePath, _
            myXlonExeSubFileCnt, myZstrExeSubFileName, myZstrExeSubFilePath, _
            myXlonOrgFileCnt, myZstrOrgFilePath, _
            myXlonSubFileCnt, myZstrSubFilePath)
    Call variablesOfxRefRunFilesDbl(myXlonExeOrgFileCnt, myZstrExeOrgFileName)   'Debug.Print
End Sub
Private Sub variablesOfxRefRunFilesDbl( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefRunFilesDbl������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim k As Long
    For k = LBound(myZvarField) To UBound(myZvarField)
        Debug.Print "�f�[�^" & k & ": " & myZvarField(k)
    Next k
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRunFilesDbl()
'//xRefRunFilesDbl���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunFilesDbl.resetConstant
End Sub
