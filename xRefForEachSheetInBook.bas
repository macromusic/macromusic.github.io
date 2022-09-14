Attribute VB_Name = "xRefForEachSheetInBook"
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���u�b�N���̑S�V�[�g���S�Z���͈́��S�}�`�ɑ΂��ď��������s����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefForEachSheetInBook"
  Private Const meMlonExeNum As Long = 0
  Private Const meMvarField As Variant = Empty
  
'//���W���[�����萔
  Private Const coXvarField As Variant = ""

'//���W���[�����萔_�񋓑�
Private Enum EnumX
'�񋓑̎g�p���̕\�L : EnumX.rowX
'��myEnum�̕\�L���[��
    '�@�V�[�gNo. : "sht" & "Enum��" & " = " & "�l" & "'�V�[�g��"
    '�A�sNo.     : "row" & "Enum��" & " = " & "�l" & "'��������V�[�g��̕�����"
    '�B��No.     : "col" & "Enum��" & " = " & "�l" & "'��������V�[�g��̕�����"
    '�C�sNo.     : "row" & "Enum��" & " = " & "�l" & "'comment" & "'��������R�����g�̕�����"
    '�D��No.     : "col" & "Enum��" & " = " & "�l" & "'comment" & "'��������R�����g�̕�����"
    shtX = 1        'Sheet1
'    rowX = 1        '�sNo
'    colX = 1        '��No
'    rowY = 1        'comment'�sNo
'    colY = 1        'comment'��No
End Enum
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXobjBook As Object

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjBook = Nothing
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
    
'//�������s
    Call callxRefForEachSheetInBook
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
    Call checkInputVariables: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables

'//�G�N�Z���u�b�N���̔C�ӂ̓���S�I�u�W�F�N�g�ɑ΂��ď��������s
    Call PabsForEachSheetInBook(myXbisExitFlag, myXobjBook)
    If myXbisExitFlag = True Then Exit Sub
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'RemP_���W���[���������ɕۑ������ϐ������o��
Private Sub remProc()
    myXbisExitFlag = False
    On Error GoTo ExitPath
    
'    If myXvarFieldIN = Empty Then myXvarFieldIN = meMvarField
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()

    Set myXobjBook = ActiveWorkbook

End Sub

 '���ۂo_�G�N�Z���u�b�N���̑S�V�[�g���S�Z���͈́��S�}�`�ɑ΂��ď��������s����
Private Sub PabsForEachSheetInBook( _
            myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object)
    myXbisExitFlag = False
  Dim myXlonShtCnt As Long: myXlonShtCnt = 0
  Dim myXobjSheet As Object
    For Each myXobjSheet In myXobjBook.Worksheets
    '//�u�b�N���̑S�V�[�g�ɑ΂��鏈��
        Call PsubPreSheetOperation(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//�V�[�g���̃f�[�^�͈͂ɑ΂��鏈��
        Call PsubForEachRangeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//�V�[�g���̑S�}�`�ɑ΂��鏈��
        Call PsubForEachShapeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//�V�[�g���̑S�O���t�ɑ΂��鏈��
        Call PsubForEachChartInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
    '//�u�b�N���̑S�V�[�g�ɑ΂��鏈��
        myXlonShtCnt = myXlonShtCnt + 1
        Call PsubPostSheetOperation(myXobjSheet)
NextPath:
    Next
    Set myXobjSheet = Nothing
    myXbisExitFlag = False
    If myXlonShtCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubPreSheetOperation(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//�u�b�N���̑S�V�[�g�ɑ΂��鏈��
'    XarbProgCode
End Sub
Private Sub PsubForEachRangeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//�V�[�g���̃f�[�^�͈͂ɑ΂��鏈��
'//�V�[�g��̃f�[�^�͈͂��擾
  Dim myXobjAllRng As Object
    With myXobjSheet
      Dim myXobjFrstRng As Object, myXobjLastRng As Object
        Set myXobjFrstRng = .Cells(1, 1)
        Set myXobjLastRng = .Cells.SpecialCells(xlCellTypeLastCell)
        Set myXobjAllRng = .Range(myXobjFrstRng, myXobjLastRng)
    End With
    Set myXobjFrstRng = Nothing: Set myXobjLastRng = Nothing
'//�f�[�^�͈͂�����
  Dim myXlonRngCnt As Long: myXlonRngCnt = 0
  Dim myXobjRange As Object
    For Each myXobjRange In myXobjAllRng
        Call PsubRangeOperation(myXbisExitFlag, myXobjRange)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonRngCnt = myXlonRngCnt + 1
NextPath:
    Next
    Set myXobjAllRng = Nothing: Set myXobjRange = Nothing
    myXbisExitFlag = False
    If myXlonRngCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubForEachShapeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//�V�[�g���̑S�}�`�ɑ΂��鏈��
  Dim myXlonShpCnt As Long: myXlonShpCnt = 0
  Dim myXobjShape As Object
    For Each myXobjShape In myXobjSheet.Shapes
        Call PsubShapeOperation(myXbisExitFlag, myXobjShape)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonShpCnt = myXlonShpCnt + 1
NextPath:
    Next
    Set myXobjShape = Nothing
    myXbisExitFlag = False
    If myXlonShpCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubForEachChartInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
    myXbisExitFlag = False
'//�V�[�g���̑S�O���t�ɑ΂��鏈��
  Dim myXlonChrtCnt As Long: myXlonChrtCnt = 0
  Dim myXobjChrtObjct As Object
    For Each myXobjChrtObjct In myXobjSheet.Charts
        Call PsubChartOperation(myXbisExitFlag, myXobjChrtObjct)
        If myXbisExitFlag = True Then GoTo NextPath
        myXlonChrtCnt = myXlonChrtCnt + 1
NextPath:
    Next
    Set myXobjChrtObjct = Nothing
    myXbisExitFlag = False
    If myXlonChrtCnt = 0 Then myXbisExitFlag = True
End Sub
Private Sub PsubRangeOperation(myXbisExitFlag As Boolean, myXobjRange As Object)
    myXbisExitFlag = False
'//�V�[�g���̃f�[�^�͈͂ɑ΂��鏈��
'    XarbProgCode
End Sub
Private Sub PsubShapeOperation(myXbisExitFlag As Boolean, myXobjShape As Object)
    myXbisExitFlag = False
'//�V�[�g���̑S�}�`�ɑ΂��鏈��
'    XarbProgCode
End Sub
Private Sub PsubChartOperation(myXbisExitFlag As Boolean, myXobjChrtObjct As Object)
    myXbisExitFlag = False
'//�V�[�g���̑S�O���t�ɑ΂��鏈��
'    XarbProgCode
End Sub
Private Sub PsubPostSheetOperation(myXobjSheet As Object)
'//�u�b�N���̑S�V�[�g�ɑ΂��鏈��
'    XarbProgCode
End Sub

'RecP_�g�p�����ϐ������W���[���������ɕۑ�����
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

'���W���[�����o_
Private Sub MsubProc()
End Sub

'���W���[�����e_
Private Function MfncFunc() As Variant
End Function

'===============================================================================================

 '��^�o_
Private Sub PfixProc()
End Sub

 '��^�e_
Private Function PfncFunc() As Variant
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

'��ModuleProc��_�G�N�Z���u�b�N���̑S�V�[�g���S�Z���͈́��S�}�`�ɑ΂��ď��������s����
Private Sub callxRefForEachSheetInBook()
  Dim myXbisCompFlag As Boolean
    Call xRefForEachSheetInBook.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefForEachSheetInBook()
'//xRefForEachSheetInBook���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefForEachSheetInBook.resetConstant
End Sub
