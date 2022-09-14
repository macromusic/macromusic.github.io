Attribute VB_Name = "xRefSlctShtData"
'Includes CSlctShtSrsData
'Includes CSlctShtDscrtData
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��͈̔͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
'Rev.002
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSlctShtData"
  Private Const meMlonExeNum As Long = 0
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
  
  Private myXlonSrsDataRowCnt As Long, myXlonSrsDataColCnt As Long, _
            myZstrShtSrsData() As String, myZvarShtSrsData() As Variant, _
            myZstrCmntData() As String
    'myZstrShtSrsData(i, j) : �擾������
    'myZvarShtSrsData(i, j) : �擾������
    'myZstrCmntData(i, j) : �擾�R�����g
  Private myXlonBgnRow As Long, myXlonEndRow As Long, _
            myXlonBgnCol As Long, myXlonEndCol As Long, _
            myXlonRows As Long, myXlonCols As Long
  
  Private myXlonDscrtDataCnt As Long, myZobjDscrtDataCell() As Object, _
            myZvarShtDscrtData() As Variant
    'myZvarShtDscrtData(i, 1) = Row
    'myZvarShtDscrtData(i, 2) = Column
    'myZvarShtDscrtData(i, 3) = SheetData
    'myZvarShtDscrtData(i, 4) = CommentData
    
'//���͐���M��
  Private myXlonSrsDataOptn As Long
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s�A���f�[�^���擾����
  
'//���̓f�[�^
  Private myXlonRngOptn As Long
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
  Private myXbisByVrnt As Boolean
    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
  Private myXbisGetCmnt As Boolean
    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
    'myXbisGetCmnt = True  : �R�����g���擾����
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean

'//���W���[�����ϐ�_�f�[�^
  Private myXstrInptBxPrmpt As String, myXstrInptBxTtl As String

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXstrInptBxPrmpt = Empty: myXstrInptBxTtl = Empty
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
    Call callxRefSlctShtData
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXobjBookOUT As Object, myXstrShtNameOUT As String, myXlonShtNoOUT As Long, _
            myXlonSrsDataRowCntOUT As Long, myXlonSrsDataColCntOUT As Long, _
            myZstrShtSrsDataOUT() As String, myZvarShtSrsDataOUT() As Variant, _
            myZstrCmntDataOUT() As String, _
            myXlonBgnRowOUT As Long, myXlonEndRowOUT As Long, _
            myXlonBgnColOUT As Long, myXlonEndColOUT As Long, _
            myXlonRowsOUT As Long, myXlonColsOUT As Long, _
            myXlonDscrtDataCntOUT As Long, _
            myZobjDscrtDataCellOUT() As Object, myZvarShtDscrtDataOUT() As Variant, _
            ByVal myXlonSrsDataOptnIN As Long, _
            ByVal myXlonRngOptnIN As Long, _
            ByVal myXbisByVrntIN As Boolean, ByVal myXbisGetCmntIN As Boolean)
    
'//���͕ϐ���������
    myXlonSrsDataOptn = Empty
    
    myXlonRngOptn = Empty
    myXbisByVrnt = False: myXbisGetCmnt = False
    
'//���͕ϐ�����荞��
    myXlonSrsDataOptn = myXlonSrsDataOptnIN
    
    myXlonRngOptn = myXlonRngOptnIN
    myXbisByVrnt = myXbisByVrntIN
    myXbisGetCmnt = myXbisGetCmntIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    Set myXobjBookOUT = Nothing: myXstrShtNameOUT = Empty: myXlonShtNoOUT = Empty

    myXlonSrsDataRowCntOUT = Empty: myXlonSrsDataColCntOUT = Empty
    Erase myZstrShtSrsDataOUT: Erase myZvarShtSrsDataOUT: Erase myZstrCmntDataOUT
    myXlonBgnRowOUT = Empty: myXlonEndRowOUT = Empty
    myXlonBgnColOUT = Empty: myXlonEndColOUT = Empty
    myXlonRowsOUT = Empty: myXlonColsOUT = Empty

    myXlonDscrtDataCntOUT = Empty
    Erase myZobjDscrtDataCellOUT: Erase myZvarShtDscrtDataOUT
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    Set myXobjBookOUT = myXobjBook
    myXstrShtNameOUT = myXstrShtName
    myXlonShtNoOUT = myXlonShtNo
    
    If myXlonSrsDataOptn = 1 Then
        myXlonSrsDataRowCntOUT = myXlonSrsDataRowCnt
        myXlonSrsDataColCntOUT = myXlonSrsDataColCnt
        myZstrShtSrsDataOUT() = myZstrShtSrsData()
        myZvarShtSrsDataOUT() = myZvarShtSrsData()
        myZstrCmntDataOUT() = myZstrCmntData()
        myXlonBgnRowOUT = myXlonBgnRow
        myXlonEndRowOUT = myXlonEndRow
        myXlonBgnColOUT = myXlonBgnCol
        myXlonEndColOUT = myXlonEndCol
        myXlonRowsOUT = myXlonRows
        myXlonColsOUT = myXlonCols
        
    ElseIf myXlonSrsDataOptn = 2 Then
        myXlonDscrtDataCntOUT = myXlonDscrtDataCnt
        myZobjDscrtDataCellOUT() = myZobjDscrtDataCell()
        myZvarShtDscrtDataOUT() = myZvarShtDscrtData()
        
    Else
    End If
    
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
    
'//S:�G�N�Z���V�[�g��̃f�[�^�͈͂�I�����ăf�[�^���擾
    Select Case myXlonSrsDataOptn
    '//�A���͈�
        Case 1
            Call setControlVariables1
            Call instCSlctShtSrsData
        
    '//�s�A���͈�
        Case 2
            Call setControlVariables2
            Call instCSlctShtDscrtData
        
    End Select
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    Set myXobjBook = Nothing: myXstrShtName = Empty: myXlonShtNo = Empty
    
    myXlonSrsDataRowCnt = Empty: myXlonSrsDataColCnt = Empty
    Erase myZstrShtSrsData: Erase myZvarShtSrsData: Erase myZstrCmntData
    myXlonBgnRow = Empty: myXlonEndRow = Empty
    myXlonBgnCol = Empty: myXlonEndCol = Empty
    myXlonRows = Empty: myXlonCols = Empty
    
    myXlonDscrtDataCnt = Empty: Erase myZobjDscrtDataCell: Erase myZvarShtDscrtData
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
    
'    If myXlonSrsDataOptn < 1 And myXlonSrsDataOptn > 2 Then GoTo ExitPath
    
'    If myXlonRngOptn < 0 And myXlonRngOptn > 3 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
    
    myXlonSrsDataOptn = 1
    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
    'myXlonSrsDataOptn = 2 : �s�A���f�[�^���擾����
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables1()
    
    myXlonRngOptn = 0
    'myXlonRngOptn = 0  : �I��͈�
    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
    'myXlonRngOptn = 3  : �S�f�[�^�͈�
    
    myXbisByVrnt = False
    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
    
    myXbisGetCmnt = True
    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
    'myXbisGetCmnt = True  : �R�����g���擾����
    
    myXstrInptBxPrmpt = ""
    myXstrInptBxTtl = ""
    
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables2()
    
    myXbisByVrnt = False
    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
    
    myXbisGetCmnt = True
    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
    'myXbisGetCmnt = True  : �R�����g���擾����
    
    myXstrInptBxPrmpt = ""
    myXstrInptBxTtl = ""
    
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

'��ClassProc��_�V�[�g��̘A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
Private Sub instCSlctShtSrsData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsSlctShtSrsData As CSlctShtSrsData: Set myXinsSlctShtSrsData = New CSlctShtSrsData
    With myXinsSlctShtSrsData
    '//�N���X���ϐ��ւ̓���
        .letRngOptn = myXlonRngOptn
        .letByVrnt = myXbisByVrnt
        .letGetCmnt = myXbisGetCmnt
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonSrsDataRowCnt = .getDataRowCnt
        myXlonSrsDataColCnt = .getDataColCnt
        If myXlonSrsDataRowCnt <= 0 Or myXlonSrsDataColCnt <= 0 Then GoTo ExitPath
        i = myXlonSrsDataRowCnt + Lo - 1: j = myXlonSrsDataColCnt + Lo - 1
        ReDim myZstrShtSrsData(i, j) As String
        ReDim myZvarShtSrsData(i, j) As Variant
        ReDim myZstrCmntData(i, j) As String
        Lc = .getOptnBase
        If myXbisByVrnt = False Then
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZstrShtSrsData(i + Lo - 1, j + Lo - 1) _
                        = .getStrShtDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        Else
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZvarShtSrsData(i + Lo - 1, j + Lo - 1) _
                        = .getVarShtDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        End If
        If myXbisGetCmnt = True Then
            For j = 1 To myXlonSrsDataColCnt
                For i = 1 To myXlonSrsDataRowCnt
                    myZstrCmntData(i + Lo - 1, j + Lo - 1) _
                        = .getCmntDataAry(i + Lc - 1, j + Lc - 1)
                Next i
            Next j
        End If
        Set myXobjBook = .getBook
        myXstrShtName = .getShtName
        myXlonShtNo = .getShtNo
        myXlonBgnRow = .getBgnEndRowCol(1, 1)
        myXlonEndRow = .getBgnEndRowCol(2, 1)
        myXlonBgnCol = .getBgnEndRowCol(1, 2)
        myXlonEndCol = .getBgnEndRowCol(2, 2)
        myXlonRows = .getBgnEndRowCol(1, 0)
        myXlonCols = .getBgnEndRowCol(0, 1)
    End With
    Set myXinsSlctShtSrsData = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSlctShtSrsData = Nothing
End Sub

'��ClassProc��_�V�[�g��̕s�A���͈͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
Private Sub instCSlctShtDscrtData()
    myXbisExitFlag = False
    
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long
  Dim myXinsSlctShtDscrtData As CSlctShtDscrtData
    Set myXinsSlctShtDscrtData = New CSlctShtDscrtData
    With myXinsSlctShtDscrtData
    '//�N���X���ϐ��ւ̓���
        .letByVrnt = myXbisByVrnt
        .letGetCmnt = myXbisGetCmnt
        .letInptBoxPrmptTtl(1) = myXstrInptBxPrmpt
        .letInptBoxPrmptTtl(2) = myXstrInptBxTtl
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXlonDscrtDataCnt = .getDataCnt
        If myXlonDscrtDataCnt <= 0 Then GoTo ExitPath
        i = myXlonDscrtDataCnt + Lo - 1
        ReDim myZobjDscrtDataCell(i) As Object
        ReDim myZvarShtDscrtData(i, Lo + 3) As Variant
        Lc = .getOptnBase
        For i = 1 To myXlonDscrtDataCnt
            Set myZobjDscrtDataCell(i + Lo - 1) = .getDataCellAry(i + Lc - 1)
        Next i
        For i = 1 To myXlonDscrtDataCnt
            myZvarShtDscrtData(i + Lo - 1, Lo + 0) = .getShtDataAry(i + Lc - 1, Lc + 0)
            myZvarShtDscrtData(i + Lo - 1, Lo + 1) = .getShtDataAry(i + Lc - 1, Lc + 1)
            myZvarShtDscrtData(i + Lo - 1, Lo + 2) = .getShtDataAry(i + Lc - 1, Lc + 2)
        Next i
        If myXbisGetCmnt = True Then
            For i = 1 To myXlonDscrtDataCnt
                myZvarShtDscrtData(i + Lo - 1, Lo + 3) = .getShtDataAry(i + Lc - 1, Lo + 3)
            Next i
        End If
        Set myXobjBook = .getBook
        myXstrShtName = .getShtName
        myXlonShtNo = .getShtNo
    End With
    Set myXinsSlctShtDscrtData = Nothing
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
    Set myXinsSlctShtDscrtData = Nothing
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
'    myXlonSrsDataOptn = 1
'    'myXlonSrsDataOptn = 1 : �A���f�[�^���擾����
'    'myXlonSrsDataOptn = 2 : �s�A���f�[�^���擾����
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables1()
'    myXlonRngOptn = 0
'    'myXlonRngOptn = 0  : �I��͈�
'    'myXlonRngOptn = 1  : �I���ʒu����ŏI�s�܂ł͈̔�
'    'myXlonRngOptn = 2  : �I���ʒu����ŏI��܂ł͈̔�
'    'myXlonRngOptn = 3  : �S�f�[�^�͈�
'    myXbisByVrnt = False
'    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
'    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
'    myXbisGetCmnt = True
'    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
'    'myXbisGetCmnt = True  : �R�����g���擾����
'    myXstrInptBxPrmpt = ""
'    myXstrInptBxTtl = ""
'End Sub
''SetP_����p�ϐ���ݒ肷��
'Private Sub setControlVariables2()
'    myXbisByVrnt = False
'    'myXbisByVrnt = False : �V�[�g�f�[�^��String�Ŏ擾����
'    'myXbisByVrnt = True  : �V�[�g�f�[�^��Variant�Ŏ擾����
'    myXbisGetCmnt = True
'    'myXbisGetCmnt = False : �R�����g���擾���Ȃ�
'    'myXbisGetCmnt = True  : �R�����g���擾����
'    myXstrInptBxPrmpt = ""
'    myXstrInptBxTtl = ""
'End Sub
'��ModuleProc��_�V�[�g��͈̔͂��w�肵�Ă��͈̔͂̃f�[�^�Ə����擾����
Private Sub callxRefSlctShtData()
'  Dim myXlonSrsDataOptn As Long, myXlonRngOptn As Long, _
'        myXbisByVrnt As Boolean, myXbisGetCmnt As Boolean
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXobjBook As Object, myXstrShtName As String, myXlonShtNo As Long
'  Dim myXlonSrsDataRowCnt As Long, myXlonSrsDataColCnt As Long, _
'        myZstrShtSrsData() As String, myZvarShtSrsData() As Variant, _
'        myZstrCmntData() As String, _
'        myXlonBgnRow As Long, myXlonEndRow As Long, _
'        myXlonBgnCol As Long, myXlonEndCol As Long, _
'        myXlonRows As Long, myXlonCols As Long
'    'myZstrShtSrsData(i, j) : �擾������
'    'myZvarShtSrsData(i, j) : �擾������
'    'myZstrCmntData(i, j) : �擾�R�����g
'  Dim myXlonDscrtDataCnt As Long, _
'        myZobjDscrtDataCell() As Object, myZvarShtDscrtData() As Variant
'    'myZvarShtDscrtData(i, 1) = Row
'    'myZvarShtDscrtData(i, 2) = Column
'    'myZvarShtDscrtData(i, 3) = SheetData
'    'myZvarShtDscrtData(i, 4) = CommentData
    Call xRefSlctShtData.callProc( _
            myXbisCmpltFlag, _
            myXobjBook, myXstrShtName, myXlonShtNo, _
            myXlonSrsDataRowCnt, myXlonSrsDataColCnt, _
            myZstrShtSrsData, myZvarShtSrsData, myZstrCmntData, _
            myXlonBgnRow, myXlonEndRow, _
            myXlonBgnCol, myXlonEndCol, _
            myXlonRows, myXlonCols, _
            myXlonDscrtDataCnt, myZobjDscrtDataCell, myZvarShtDscrtData, _
            myXlonSrsDataOptn, myXlonRngOptn, myXbisByVrnt, myXbisGetCmnt)
    Call variablesOfxRefSlctShtData(myXlonSrsDataRowCnt, myZstrShtSrsData)   'Debug.Print
'    Call variablesOfxRefSlctShtData(myXlonDscrtDataCnt, myZvarShtDscrtData)  'Debug.Print
End Sub
Private Sub variablesOfxRefSlctShtData( _
            ByVal myXlonDataCnt As Long, ByRef myZvarField As Variant)
'//xRefSlctShtData������o�͂����ϐ��̓��e�m�F
    Debug.Print "�f�[�^��: " & myXlonDataCnt
    If myXlonDataCnt <= 0 Then Exit Sub
  Dim i As Long, j As Long
    For i = LBound(myZvarField, 1) To UBound(myZvarField, 1)
        For j = LBound(myZvarField, 2) To UBound(myZvarField, 2)
            Debug.Print "�f�[�^" & i & "," & j & ": " & myZvarField(i, j)
        Next j
    Next i
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefSlctShtData()
'//xRefSlctShtData���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSlctShtData.resetConstant
End Sub
