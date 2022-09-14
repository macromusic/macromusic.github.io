Attribute VB_Name = "xRefShtDataByEnm"
'Includes PfixGetSheetRangeDataVariant
'Includes PfncbisCheckArrayDimension
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�V�[�g��̕�����ʒu��񋓑̂Ŏw�肵�ăf�[�^���擾����
'Rev.002
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefShtDataByEnm"
  Private Const meMlonExeNum As Long = 0

'//���W���[�����萔_�񋓑�
Public Enum EnumX
'�񋓑̎g�p���̕\�L : EnumX.rowX

'//[Sheet1]�V�[�g��̃p�����[�^�z�u���`
    shtX = 2                        'Sheet1

    rowFldrPth = 4                  '���t�H���_�p�X �F
    rowFileExt = 5                  '���t�@�C���g���q �F
    colData = 3                     'comment'�f�[�^��

    rowData = 8                     'comment'�f�[�^�s
    colFilePth = 2                  '���t�@�C���ꗗ
End Enum
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^_1
  Private myXstrFldrPth As String       '
  Private myXstrFileExt As String       '
  
'//�o�̓f�[�^_2
  Private myZstrFilePth() As String     '
  
'//�o�̓f�[�^_3
  Private myXobjDataRng As Object   '
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData As Variant

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonRowCnt = Empty: myXlonColCnt = Empty: myZvarShtData = Empty
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
    Call callxRefShtDataByEnm
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXstrFldrPthOUT As String, myXstrFileExtOUT As String, _
            myZstrFilePthOUT() As String, _
            myXobjDataRngOUT As Object)
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    myXstrFldrPthOUT = Empty: myXstrFileExtOUT = Empty
    Erase myZstrFilePthOUT
    Set myXobjDataRngOUT = Nothing
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXstrFldrPthOUT = myXstrFldrPth
    myXstrFileExtOUT = myXstrFileExt
    
  Dim k As Long
    k = UBound(myZstrFilePth)
    ReDim myZstrFilePthOUT(k) As String
    For k = LBound(myZstrFilePth) To UBound(myZstrFilePth)
        myZstrFilePthOUT(k) = myZstrFilePth(k)
    Next k
    
    Set myXobjDataRngOUT = myXobjDataRng

End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//S:�V�[�g��̑S�f�[�^���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:�K�v�ȏ��𒊏o
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    
    myXstrFldrPth = Empty: myXstrFileExt = Empty
    
    Erase myZstrFilePth

    Set myXobjDataRng = Nothing
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
End Sub

'SnsP_�V�[�g��̑S�f�[�^���擾
Private Sub snsProc()
    myXbisExitFlag = False
    
'//�V�[�g��̎w��͈͂܂��̓f�[�^�S�͈͂̃f�[�^��Variant�ϐ��Ɏ捞��
  Dim i As Long: i = EnumX.shtX
  Dim myXobjSheet As Object, myXobjFrstCell As Object, myXobjLastCell As Object
    Set myXobjSheet = ThisWorkbook.Worksheets(i)
    
    Call PfixGetSheetRangeDataVariant( _
            myXlonRowCnt, myXlonColCnt, myZvarShtData, _
            myXobjSheet, myXobjFrstCell, myXobjLastCell)
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then GoTo ExitPath
    
    Set myXobjSheet = Nothing
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_�K�v�ȏ��𒊏o
Private Sub prsProc()
    myXbisExitFlag = False
    
  Dim i As Long, j As Long

    On Error Resume Next
    
'//1
    j = EnumX.colData
    
    i = EnumX.rowFldrPth
    myXstrFldrPth = CStr(myZvarShtData(i, j))
    
    i = EnumX.rowFileExt
    myXstrFileExt = CStr(myZvarShtData(i, j))
    
'//2
    i = EnumX.rowData
    j = EnumX.colFilePth
    
  Dim myXstrTmp As String, k As Long, n As Long: n = 0
    For k = i To UBound(myZvarShtData, 1)
        myXstrTmp = Empty
        myXstrTmp = CStr(myZvarShtData(k, j))
        If myXstrTmp = "" Then Exit For
        
        n = n + 1: ReDim Preserve myZstrFilePth(n) As String
        myZstrFilePth(n) = myXstrTmp
    Next k
    
'//3
  Dim myXobjSheet As Object
    k = EnumX.shtX
    Set myXobjSheet = ThisWorkbook.Worksheets(k)
    
    i = EnumX.rowData
    j = EnumX.colFilePth
    Set myXobjDataRng = myXobjSheet.Cells(i, j)
    
'  Dim rb As Long, cb As Long, re As Long, ce As Long
'    rb = EnumX.rowFldrPth
'    cb = EnumX.colData
'    re = EnumX.rowFileExt
'    ce = EnumX.colData
'    With myXobjSheet
'        Set myXobjDataRng = .Range(.Cells(rb, cb), .Cells(re, ce))
'    End With
    
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'checkP_�o�͕ϐ����e���m�F����
Private Sub checkOutputVariables()
    myXbisExitFlag = False
    
    If myXstrFldrPth = "" Then GoTo ExitPath
    If myXstrFileExt = "" Then GoTo ExitPath
    
    If PfncbisCheckArrayDimension(myZstrFilePth, 1) = False Then GoTo ExitPath
    
    If myXobjDataRng Is Nothing Then GoTo ExitPath
    
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

 '��^�o_�V�[�g��̎w��͈͂܂��̓f�[�^�S�͈͂̃f�[�^��Variant�ϐ��Ɏ捞��
Private Sub PfixGetSheetRangeDataVariant( _
            myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData As Variant, _
            ByVal myXobjSheet As Object, _
            ByVal myXobjFrstCell As Object, ByVal myXobjLastCell As Object)
'myZvarShtData(i, j) : �f�[�^
    myXlonRowCnt = Empty: myXlonColCnt = Empty: myZvarShtData = Empty
    If myXobjSheet Is Nothing Then Exit Sub
'//�V�[�g��̎w��͈͂��I�u�W�F�N�g�z��Ɏ捞��
  Dim myXobjShtRng As Object
    If myXobjFrstCell Is Nothing Then Set myXobjFrstCell = myXobjSheet.Cells(1, 1)
    If myXobjLastCell Is Nothing Then _
        Set myXobjLastCell = myXobjSheet.Cells.SpecialCells(xlCellTypeLastCell)
    Set myXobjShtRng = myXobjSheet.Range(myXobjFrstCell, myXobjLastCell)
    myXlonRowCnt = myXobjShtRng.Rows.Count
    myXlonColCnt = myXobjShtRng.Columns.Count
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then Exit Sub
'//�I�u�W�F�N�g�z�񂩂�f�[�^���擾
    myZvarShtData = myXobjShtRng.Value
    Set myXobjShtRng = Nothing
End Sub

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

'��ModuleProc��_�V�[�g��̕�����ʒu��񋓑̂Ŏw�肵�ăf�[�^���擾����
Private Sub callxRefShtDataByEnm()
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXstrFldrPth As String, myXstrFileExt As String
'  Dim myZstrFilePth() As String
'  Dim myXobjDataRng As Object   '
    Call xRefShtDataByEnm.callProc( _
            myXbisCmpltFlag, myXstrFldrPth, myXstrFileExt, myZstrFilePth, myXobjDataRng)
    Debug.Print myXbisCmpltFlag
    Debug.Print myXstrFldrPth
  Dim k As Long
    For k = LBound(myZstrFilePth) To UBound(myZstrFilePth)
        Debug.Print "�f�[�^" & k & ": " & myZstrFilePth(k)
    Next k
    Debug.Print myXobjDataRng.Value
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetxRefShtDataByEnm()
'//xRefShtDataByEnm���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefShtDataByEnm.resetConstant
End Sub
