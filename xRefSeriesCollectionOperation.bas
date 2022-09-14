Attribute VB_Name = "xRefSeriesCollectionOperation"
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���O���t���̑S�n��ɑ΂��ď��������s����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefSeriesCollectionOperation"
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
  Private myXobjChart As Object
  Private myXlonSrsClctnCnt As Long, myXlonErrSrsClctnCnt As Long, myZobjErrSrsClctn() As Object

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    Set myXobjChart = Nothing
    myXlonSrsClctnCnt = Empty
    myXlonErrSrsClctnCnt = Empty
    Erase myZobjErrSrsClctn
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
    Call callxRefSeriesCollectionOperation
    
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

'//�G�N�Z���O���t���̑S�n��ɑ΂��ď��������s
    Call PabsSeriesCollectionOperation( _
            myXbisExitFlag, myXlonSrsClctnCnt, myXlonErrSrsClctnCnt, myZobjErrSrsClctn, _
            myXobjChart)
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

    Set myXobjOrg = Selection

End Sub

 '���ۂo_�G�N�Z���O���t���̑S�n��ɑ΂��ď��������s����
Private Sub PabsSeriesCollectionOperation( _
            myXbisExitFlag As Boolean, myXlonSrsClctnCnt As Long, _
            myXlonErrSrsClctnCnt As Long, myZobjErrSrsClctn() As Object, _
            ByVal myXobjChart As Object)
    myXlonSrsClctnCnt = Empty: myXlonErrSrsClctnCnt = Empty: Erase myZobjErrSrsClctn
    On Error GoTo ExitPath
  Dim k As Long: k = Chart.SeriesCollection.Count
    If k <= 0 Then GoTo ExitPath
    On Error GoTo 0
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim myXobjSrsClctn As Object, n As Long, e As Long: n = 0: e = Lo - 1
    For Each myXobjSrsClctn In myXobjChart.SeriesCollection
        Call PsubSeriesCollectionOperation(myXbisExitFlag, myXobjSrsClctn)
        If myXbisExitFlag = True Then
            e = e + 1: ReDim Preserve myZobjErrSrsClctn(e) As Object
            Set myZobjErrSrsClctn(e) = myXobjSrsClctn
        Else
            n = n + 1
        End If
    Next myXobjSrsClctn
    myXlonSrsClctnCnt = n: myXlonErrSrsClctnCnt = e - Lo + 1
    If myXlonErrSrsClctnCnt >= 1 Then
        myXbisExitFlag = True
    Else
        myXbisExitFlag = False
    End If
    Set myXobjSrsClctn = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub
Private Sub PsubSeriesCollectionOperation(myXbisExitFlag As Boolean, _
            ByVal myXobjSrsClctn As Object)
    myXbisExitFlag = False
'//�V�[�g���̑S�}�`�ɑ΂��鏈��
'    XarbProgCode
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

'��ModuleProc��_�G�N�Z���O���t���̑S�n��ɑ΂��ď��������s����
Private Sub callxRefSeriesCollectionOperation()
  Dim myXbisCompFlag As Boolean
    Call xRefSeriesCollectionOperation.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefSeriesCollectionOperation()
'//xRefSeriesCollectionOperation���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefSeriesCollectionOperation.resetConstant
End Sub
