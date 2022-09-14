Attribute VB_Name = "xRefRunInfos"
'Includes PfncbisCheckArrayDimension
'Includes PfncbisCheckArrayDimensionLength
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�������ɑ΂��ĘA�����������{����
'Rev.002
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRunInfos"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXlonExeInfoCnt As Long, myZstrExeInfo() As String
    'myZstrExeInfo(i) : ���s���
  
'//���͐���M��
  
'//���̓f�[�^
  Private myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
    'myZstrOrgInfo(i) : �����
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonRunInfoNo As Long, myXstrRunInfo As String

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRunInfoNo = Empty: myXstrRunInfo = Empty
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

'    myXlonOrgInfoCnt = 1
'    ReDim myZstrOrgInfo(myXlonOrgInfoCnt) As String
    
'//�������s
    Call callxRefRunInfos
    
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXlonExeInfoCntOUT As Long, myZstrExeInfoOUT() As String, _
            ByVal myXlonOrgInfoCntIN As Long, ByRef myZstrOrgInfoIN() As String)
    
'//���͕ϐ���������
    myXlonOrgInfoCnt = Empty
    Erase myZstrOrgInfo

'//���͕ϐ�����荞��
    If myXlonOrgInfoCntIN <= 0 Then Exit Sub
    myXlonOrgInfoCnt = myXlonOrgInfoCntIN
    myZstrOrgInfo() = myZstrOrgInfoIN()
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    myXlonExeInfoCntOUT = Empty
    Erase myZstrExeInfoOUT
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    If myXlonExeInfoCnt <= 0 Then Exit Sub
    myXlonExeInfoCntOUT = myXlonExeInfoCnt
    myZstrExeInfoOUT() = myZstrExeInfo()
    
ExitPath:
End Sub

'PublicF_
Public Function fncbisCmpltFlag( _
            ByVal myXlonOrgInfoCntIN As Long, _
            ByRef myZstrOrgInfoIN() As String) As Boolean
    fncbisCmpltFlag = Empty
    
'//���͕ϐ���������
    myXlonOrgInfoCnt = Empty
    Erase myZstrOrgInfo

'//���͕ϐ�����荞��
    If myXlonOrgInfoCntIN <= 0 Then Exit Sub
    myXlonOrgInfoCnt = myXlonOrgInfoCntIN
    myZstrOrgInfo() = myZstrOrgInfoIN()
    
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
    
'//C:���ꗗ���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
 
    '//S:�e���̃f�[�^�擾����
        Call snsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "4-" & k   'PassFlag
 
    '//P:�e���̃f�[�^���H����
        Call prsProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "5-" & k   'PassFlag
            
    '//Run:�e���̃f�[�^�o�͏���
        Call runProcForLoop
        If myXbisExitFlag = True Then GoTo NextPath
'        Debug.Print "PassFlag: " & meMstrMdlName & "6-" & k   'PassFlag
        
        n = n + 1
        ReDim Preserve myZstrExeInfo(n) As String
        myZstrExeInfo(n) = myXstrRunInfo
NextPath:
    Next k
    myXlonExeInfoCnt = n - Lo + 1
'    Debug.Print "PassFlag: " & meMstrMdlName & "7"    'PassFlag
    
'//P:Loop��̉��H����
    Call prsProcAfterLoop
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "8"    'PassFlag

'//Run:�t�@�C�i���C�Y����
    Call runFinalize
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "9"   'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXlonExeInfoCnt = Empty: Erase myZstrExeInfo
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
    
'//�z��ϐ��̎��������擾
  Dim myXlonAryDmnsn As Long
    myXlonAryDmnsn = PfnclonArrayDimension(myZstrOrgInfo)
    
  Dim Li As Long, myXstrTmp As String
  
    On Error GoTo ExitPath
    Select Case myXlonAryDmnsn
        Case 1
            Li = LBound(myZstrOrgInfo): myXstrTmp = myZstrOrgInfo(Li)
            Exit Sub
        Case 2
            Li = LBound(myZstrOrgInfo, 1): myXstrTmp = myZstrOrgInfo(Li, Li)
        Case Else: GoTo ExitPath
    End Select
    On Error GoTo 0

  Dim myZstrOrgInfoINT() As String
    myZstrOrgInfoINT() = myZstrOrgInfo()
    
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
    i = UBound(myZstrOrgInfoINT, 1)
    ReDim myZstrOrgInfo(i) As String
    For i = LBound(myZstrOrgInfoINT) To UBound(myZstrOrgInfoINT)
        myZstrOrgInfo(i + Lo - Li) = myZstrOrgInfoINT(i)
    Next i
    
    Erase myZstrOrgInfoINT
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'CtrlP_
Private Sub ctrRunFiles()

'//C:���ꗗ���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim myXvarTmpInfo As Variant, k As Long: k = Li - 1
    For Each myXvarTmpPath In myZstrOrgInfo
        myXstrRunInfo = Empty
        k = k + 1: myXlonRunInfoNo = k
        myXstrRunInfo = CStr(myXvarTmpInfo)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next myXvarTmpInfo
    myXlonExeInfoCnt = n - Lo + 1
    
'//C:���ꗗ���������s
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim n As Long: n = Lo - 1
  Dim Li As Long: Li = LBound(myZstrOrgInfo)
  Dim k As Long
    For k = LBound(myZstrOrgInfo) To UBound(myZstrOrgInfo)
        myXstrRunInfo = Empty
        myXlonRunInfoNo = k
        myXstrRunInfo = myZstrOrgInfo(k)
        If myXstrRunInfo = "" Then GoTo NextPath
        'XarbProgCode
NextPath:
    Next k
    myXlonExeInfoCnt = n - Lo + 1
    
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

'SnsP_�e���̃f�[�^�擾����
Private Sub snsProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�e���̃f�[�^���H����
Private Sub prsProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�e���̃f�[�^�o�͏���
Private Sub runProcForLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "6-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_Loop��̉��H����
Private Sub prsProcAfterLoop()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "8-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'RunP_�t�@�C�i���C�Y����
Private Sub runFinalize()
    myXbisExitFlag = False
    
'    If myXvarField = "" Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "9-1"     'PassFlag
    
    'XarbProgCode
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_�o�͕ϐ����e���m�F����
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    If myXlonExeInfoCnt <= 0 Then GoTo ExitPath
    If PfncbisCheckArrayDimension(myZstrExeInfo, 1) = False Then GoTo ExitPath
    
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

 '��^�e_�z��ϐ��̎��������擾����
Private Function PfnclonArrayDimension(ByRef myZvarOrgData As Variant) As Long
    PfnclonArrayDimension = Empty
    If IsArray(myZvarOrgData) = False Then Exit Function
  Dim myXvarTmp As Variant, k As Long: k = 0
    On Error Resume Next
    Do
        k = k + 1: myXvarTmp = UBound(myZvarOrgData, k)
    Loop While Err.Number = 0
    On Error GoTo 0
    PfnclonArrayDimension = k - 1
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

'��ModuleProc��_�������ɑ΂��ĘA�����������{����
Private Sub callxRefRunInfos()
'  Dim myXlonOrgInfoCnt As Long, myZstrOrgInfo() As String
'    'myZstrOrgInfo(i) : �����
'    myXlonOrgInfoCnt = XarbLong
'    ReDim myZstrOrgInfo(1) As String
'    myZstrOrgInfo(1) = XarbString
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonExeInfoCnt As Long, myZstrExeInfo() As String
'    'myZstrExeInfo(i) : ���s���
    Call xRefRunInfos.callProc( _
            myXbisCmpltFlag, myXlonExeInfoCnt, myZstrExeInfo, _
            myXlonOrgInfoCnt, myZstrOrgInfo)
    Call variablesOfxRefRunInfos(myXlonExeInfoCnt, myZstrExeInfo)    'Debug.Print
End Sub
Private Sub variablesOfxRefRunInfos( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefRunInfos������o�͂����ϐ��̓��e�m�F
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
Private Sub resetConstantInxRefRunInfos()
'//xRefRunInfos���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRunInfos.resetConstant
End Sub
