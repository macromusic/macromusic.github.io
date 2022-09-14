Attribute VB_Name = "xRefExtrctData"
'Includes CSeriesData
'Includes CVrblToSht
'Includes PfixGetSheetRangeDataVariant
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_���f�[�^����K�v�ȃf�[�^���擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefExtrctData"
  Private Const meMlonExeNum As Long = 0

'//���W���[�����萔_�񋓑�
Private Enum EnumX
'�񋓑̎g�p���̕\�L : EnumX.rowX
'��myEnum�̕\�L���[��
    '�@�V�[�gNo. : "sht" & "Enum��" & " = " & "�l" & "'�V�[�g��"
    '�A�sNo.     : "row" & "Enum��" & " = " & "�l" & "'��������V�[�g��̕�����"
    '�B��No.     : "col" & "Enum��" & " = " & "�l" & "'��������V�[�g��̕�����"
    '�C�sNo.     : "row" & "Enum��" & " = " & "�l" & "'comment" & "'��������R�����g�̕�����"
    '�D��No.     : "col" & "Enum��" & " = " & "�l" & "'comment" & "'��������R�����g�̕�����"
    
    shtExe1 = 1         '����
    rowBgn = 1          '�R�[�h
    colBgn = 1          '�R�[�h
    rowPst = 1          '�����R�[�h
    ColPst = 4          '�����R�[�h
    RowOfset1 = 1
    
    shtExe2 = 2         '���f�[�^
    rowOrg = 3          '�����R�[�h
    colOrg = 2          '�����R�[�h
    RowOfset2 = 1
    
End Enum
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^
  Private myXlonSrsDataCnt As Long, myZstrSrsData() As String
    'myZstrSrsData(k) : �擾������
  Private myXbisRowDrctn As Boolean
    'myXbisRowDrctn = True  : �s�����݂̂�����
    'myXbisRowDrctn = False : ������݂̂�����
  Private myXlonExtShtNo As Long, myXobjExtSheet As Object
  Private myXlonBgnRow As Long, myXlonBgnCol As Long

  Private myXobjOrgSheet As Object, myXobjFrstCell As Object, myXobjLastCell As Object
  Private myXlonRowCnt As Long, myXlonColCnt As Long, myZvarShtData As Variant
    'myZvarShtData(i, j) : �V�[�g�f�[�^

  Private myXlonArngDataCnt As Long, myZvarArngData() As Variant

  Private myXobjPstFrstCell As Object

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
    myXlonSrsDataCnt = Empty: Erase myZstrSrsData
    myXbisRowDrctn = False
    myXlonExtShtNo = Empty: Set myXobjExtSheet = Nothing
    myXlonBgnRow = Empty: myXlonBgnCol = Empty
    Set myXobjOrgSheet = Nothing
    Set myXobjFrstCell = Nothing: Set myXobjLastCell = Nothing
    myXlonRowCnt = Empty: myXlonColCnt = Empty: myZvarShtData = Empty
    myXlonArngDataCnt = Empty: Erase myZvarArngData
    Set myXobjPstFrstCell = Nothing
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
    Call callxRefExtrctData
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc(myXbisCmpltFlagOUT As Boolean)
Application.ScreenUpdating = False
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
Application.ScreenUpdating = True
End Sub

'CtrlP_
Private Sub ctrProc()
    Call initializeOutputVariables
    Call initializeModuleVariables
    Call remProc: If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
'//C:����p�ϐ���ݒ�
    Call setControlVariables
    
'//S:�V�[�g��̘A������f�[�^�͈͂��擾
    Call instCSeriesData
    If myXlonSrsDataCnt <= 0 Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//S:�V�[�g��̎w��͈͂܂��̓f�[�^�S�͈͂̃f�[�^��Variant�ϐ��Ɏ捞��
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag
    
'//P:���f�[�^����K�v�ȃf�[�^�𒊏o
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//Run:�ϐ������G�N�Z���V�[�g�ɏ����o��
    Call instCVrblToSht
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
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

    myXbisRowDrctn = True
    
    myXlonExtShtNo = EnumX.shtExe1
    Set myXobjExtSheet = ThisWorkbook.Worksheets(myXlonExtShtNo)
    
    myXlonBgnRow = EnumX.rowBgn + EnumX.RowOfset1
    myXlonBgnCol = EnumX.colBgn
    
  Dim myXlonPstRow As Long, myXlonPstCol As Long
    myXlonPstRow = EnumX.rowPst + EnumX.RowOfset1
    myXlonPstCol = EnumX.ColPst
    Set myXobjPstFrstCell = ThisWorkbook.Worksheets(myXlonExtShtNo)
    Set myXobjPstFrstCell = myXobjPstFrstCell.Cells(myXlonPstRow, myXlonPstCol)
    
  Dim myXlonOrgShtNo As Long, myXlonFrstRow As Long, myXlonFrstCol As Long
    myXlonOrgShtNo = EnumX.shtExe2
    Set myXobjOrgSheet = ThisWorkbook.Worksheets(myXlonOrgShtNo)
    myXlonFrstRow = EnumX.rowOrg + EnumX.RowOfset2
    myXlonFrstCol = EnumX.colOrg
    Set myXobjFrstCell = myXobjOrgSheet.Cells(myXlonFrstRow, myXlonFrstCol)
    
End Sub

'SnsP_�V�[�g��̎w��͈͂܂��̓f�[�^�S�͈͂̃f�[�^��Variant�ϐ��Ɏ捞��
Private Sub snsProc()
    myXbisExitFlag = False
    
    Call PfixGetSheetRangeDataVariant( _
            myXlonRowCnt, myXlonColCnt, myZvarShtData, _
            myXobjOrgSheet, myXobjFrstCell, myXobjLastCell)
    If myXlonRowCnt <= 0 Or myXlonColCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_���f�[�^����K�v�ȃf�[�^�𒊏o����
Private Sub prsProc()
    myXbisExitFlag = False
    
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
  Dim k As Long, i As Long, j As Long
  Dim myXstrOrgData As String, myXstrExtCode As String
    
    k = myXlonSrsDataCnt + L - 1
    j = myXlonColCnt + L - 1
    ReDim myZvarArngData(k, j) As Variant
    
    For k = LBound(myZstrSrsData) To UBound(myZstrSrsData)
        myXstrOrgData = myZstrSrsData(k)
        
        For i = LBound(myZvarShtData, 1) To UBound(myZvarShtData, 1)
            myXstrExtCode = myZvarShtData(i, L)
            
            If myXstrExtCode = myXstrOrgData Then
                myXlonArngDataCnt = myXlonArngDataCnt + 1
                For j = LBound(myZvarShtData, 2) To UBound(myZvarShtData, 2)
                    myZvarArngData(k, j) = myZvarShtData(i, j)
                Next j
            End If
            
        Next i
        
    Next k
    If myXlonArngDataCnt >= myXlonSrsDataCnt Then GoTo ExitPath
    
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

'��ClassProc��_�V�[�g��̘A������f�[�^�͈͂��擾����
Private Sub instCSeriesData()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim k As Long
  Dim myXinsSeriesData As CSeriesData: Set myXinsSeriesData = New CSeriesData
    With myXinsSeriesData
    '//�N���X���ϐ��ւ̓���
        Set .setSrchSheet = myXobjExtSheet
        .letBgnRowCol(1) = myXlonBgnRow
        .letBgnRowCol(2) = myXlonBgnCol
        .letRowDrctn = True
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

'��ClassProc��_�ϐ������G�N�Z���V�[�g�ɏ����o��
Private Sub instCVrblToSht()
  Dim myXinsVrblToSht As CVrblToSht: Set myXinsVrblToSht = New CVrblToSht
    With myXinsVrblToSht
    '//�N���X���ϐ��ւ̓���
        .letVrbl = myZvarArngData
        Set .setPstFrstCell = myXobjPstFrstCell
        .letInptBxOFF = False
        .letEachWrtON = False
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        myXbisExitFlag = Not .fncbisCmpltFlag
    End With
    Set myXinsVrblToSht = Nothing
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

'��ModuleProc��_����ƍ������Z�f�[�^����荞��
Private Sub callxRefExtrctData()
  Dim myXbisCmpltFlag As Boolean
    Call xRefExtrctData.callProc(myXbisCmpltFlag)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefExtrctData()
'//xRefExtrctData���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefExtrctData.resetConstant
End Sub
