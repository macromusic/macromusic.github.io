Attribute VB_Name = "MexeHypLnkRltvToAbslt"
'Includes PincGetExcelBookObject
'Includes PfncstrGetHyperLinkBase
'Includes PfixSetHyperLinkBase
'Includes PabsForEachSheetInBook
'Includes PfncstrGetHyperLinkPathAtRange
'Includes PfncstrGetAbsolutePath
'Includes PfncbisCheckFolderExist
'Includes PfncbisCheckFileExist
'Includes PfixSetHyperLinkWithSheetCellAtRange
'Includes PfncstrGetHyperLinkPathAtShape
'Includes PfixSetHyperLinkAtShape
'Includes PfixOverwriteSaveExcelBook
'Includes PfixCloseExcelBook

Option Explicit
Option Base 1

'��ModuleProc��_�G�N�Z���u�b�N���̑S�n�C�p�[�����N���ΎQ�ƂɕύX����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "MexeHypLnkRltvToAbslt"

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
    rowX = 1        '�sNo
    colX = 1        '��No
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

'PublicP_
Public Sub exeProc()
    
'//�������s
    Call callMexeHypLnkToAbslt
    
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
    
'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
End Sub

'CtrlP_
Private Sub ctrProc()
    myXbisCmpltFlag = False
    Call initializeModuleVariables
'    Debug.Print "PassFlag: " & meMstrMdlName & "1"     'PassFlag
    
  Dim myXstrFullName As String
    myXstrFullName = "C:\Users\Hiroki\Documents\_VBA4XPC\11_�v���O�����f�[�^�x�[�X\01_VBA�\��\c10_�n�C�p�[�����N" _
        & "\" & "test.xlsm"
'    myXstrFullName = ThisWorkbook.Worksheets(EnumX.shtX).Cells(EnumX.rowX, EnumX.colX).Value
    
'//�w��G�N�Z���u�b�N�̏�Ԃ��m�F���ăI�u�W�F�N�g���擾
    Call PincGetExcelBookObject(myXbisExitFlag, myXobjBook, myXstrFullName)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "�f�[�^: " & myXobjBook.Name

'//�G�N�Z���u�b�N�̃v���p�e�B�̃n�C�p�[�����N�̊�_���擾
  Dim myXstrHypLnkBase As String
    myXstrHypLnkBase = PfncstrGetHyperLinkBase(myXobjBook)
    If myXstrHypLnkBase <> "" Then GoTo ExitPath

'//�G�N�Z���u�b�N�̃v���p�e�B�̃n�C�p�[�����N�̊�_��ݒ�
    myXstrHypLnkBase = "*"
    Call PfixSetHyperLinkBase(myXbisExitFlag, myXobjBook, myXstrHypLnkBase)
    If myXbisExitFlag = True Then GoTo ExitPath

'//�G�N�Z���u�b�N���̔C�ӂ̓���S�I�u�W�F�N�g�ɑ΂��ď��������s
    Call PabsForEachSheetInBook(myXbisExitFlag, myXobjBook)
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "2"     'PassFlag

'//�G�N�Z���u�b�N���㏑���ۑ�
  Dim myXstrBookName As String
    myXstrBookName = myXobjBook.Name
    Call PfixOverwriteSaveExcelBook(myXbisExitFlag, myXstrBookName)
    If myXbisExitFlag = True Then GoTo ExitPath

'//�G�N�Z���u�b�N�����
    Call PfixCloseExcelBook(myXbisExitFlag, myXstrBookName)
    If myXbisExitFlag = True Then GoTo ExitPath
    
    myXbisCmpltFlag = True
ExitPath:
    Call initializeModuleVariables
End Sub

 '���ۂo_�G�N�Z���u�b�N���̑S�V�[�g���S�Z���͈́��S�}�`�ɑ΂��ď��������s����
Private Sub PabsForEachSheetInBook( _
            myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object)
    myXbisExitFlag = False
  Dim myXlonShtCnt As Long: myXlonShtCnt = 0
  Dim myXobjSheet As Object
    For Each myXobjSheet In myXobjBook.Worksheets
        myXlonShtCnt = myXlonShtCnt + 1
    '//�V�[�g���̃f�[�^�͈͂ɑ΂��鏈��
        Call PsubForEachRangeInSheet(myXbisExitFlag, myXobjSheet)
        If myXbisExitFlag = True Then GoTo NextPath
'    '//�V�[�g���̑S�}�`�ɑ΂��鏈��
'        Call PsubForEachShapeInSheet(myXbisExitFlag, myXobjSheet)
'        If myXbisExitFlag = True Then GoTo NextPath
NextPath:
    Next
    Set myXobjSheet = Nothing
    myXbisExitFlag = False
    If myXlonShtCnt = 0 Then myXbisExitFlag = True
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
'Private Sub PsubForEachShapeInSheet(myXbisExitFlag As Boolean, myXobjSheet As Object)
'    myXbisExitFlag = False
''//�V�[�g���̑S�}�`�ɑ΂��鏈��
'  Dim myXlonShpCnt As Long: myXlonShpCnt = 0
'  Dim myXobjShape As Object
'    For Each myXobjShape In myXobjSheet.Shapes
'        Call PsubShapeOperation(myXbisExitFlag, myXobjShape)
'        If myXbisExitFlag = True Then GoTo NextPath
'        myXlonShpCnt = myXlonShpCnt + 1
'NextPath:
'    Next
'    Set myXobjShape = Nothing
'    myXbisExitFlag = False
'    If myXlonShpCnt = 0 Then myXbisExitFlag = True
'End Sub
Private Sub PsubRangeOperation(myXbisExitFlag As Boolean, myXobjRange As Object)
    myXbisExitFlag = False
'//�V�[�g���̃f�[�^�͈͂ɑ΂��鏈��
    
'//�w��Z���͈͂ɐݒ肳�ꂽ�n�C�p�[�����N��̃p�X���擾
  Dim myXstrLinkPath As String
    myXstrLinkPath = PfncstrGetHyperLinkPathAtRange(myXobjRange)
    If myXstrLinkPath = "" Then Exit Sub
    
'//���΃t�@�C���p�X���w�肵�Đ�΃p�X���擾
  Dim myXstrRltvPath As String
    myXstrRltvPath = myXstrLinkPath
  Dim myXstrAbsltPath As String
    myXstrAbsltPath = PfncstrGetAbsolutePath(myXstrRltvPath, myXobjBook)

'//�w��t�H���_�̑��݂��m�F
  Dim myXbisFldrExistFlag As Boolean
    myXbisFldrExistFlag = PfncbisCheckFolderExist(myXstrAbsltPath)

'//�w��t�@�C���̑��݂��m�F
  Dim myXbisFileExistFlag As Boolean
    myXbisFileExistFlag = PfncbisCheckFileExist(myXstrAbsltPath)
    
    If myXbisFldrExistFlag = False And myXbisFldrExistFlag = False Then
        Debug.Print "�p�X�G���[: " & myXstrRltvPath
        myXbisExitFlag = True
        Exit Sub
    End If

'//�w��Z���͈͂Ƀn�C�p�[�����N��ݒ�
  Dim myXstrHypLnkAdrs As String, myXstrSubAdrs As String, myXstrTxt As String
    myXstrHypLnkAdrs = myXstrAbsltPath
    myXstrSubAdrs = ""
    myXstrTxt = ""
    Call PfixSetHyperLinkWithSheetCellAtRange( _
            myXbisExitFlag, _
            myXobjRange, myXstrHypLnkAdrs, myXstrSubAdrs, myXstrTxt)
    If myXbisExitFlag = True Then Exit Sub

End Sub
'Private Sub PsubShapeOperation(myXbisExitFlag As Boolean, myXobjShape As Object)
'    myXbisExitFlag = False
''//�V�[�g���̑S�}�`�ɑ΂��鏈��
''    XarbProgCode
'End Sub
'End Sub

'===============================================================================================

 '��^�o_�w��G�N�Z���u�b�N�̏�Ԃ��m�F���ăI�u�W�F�N�g���擾����
Private Sub PincGetExcelBookObject( _
            myXbisExitFlag As Boolean, myXobjBook As Object, _
            ByVal myXstrFullName As String)
'Includes PfnclonCheckExcelBookOpening
'Includes PfncobjOpenExcelBook
'Includes PfncobjGetFile
'Includes PfixCloseExcelBook
'Includes PfncobjGetExcelBookIfOpened
    myXbisExitFlag = False: Set myXobjBook = Nothing
  Dim myXlonCheckBookOpening As Long, myXstrBookName As String
'//�w��G�N�Z���u�b�N�����ɊJ���Ă��邩�m�F
    myXlonCheckBookOpening = PfnclonCheckExcelBookOpening(myXstrFullName)
    Select Case myXlonCheckBookOpening
        Case 0
            myXbisExitFlag = True
            Exit Sub
        Case 1
        '//�t�@�C���p�X���w�肵�ăG�N�Z���u�b�N���J��
            Set myXobjBook = PfncobjOpenExcelBook(myXstrFullName)
        Case 2
        '//�w��t�@�C���̃I�u�W�F�N�g���擾
            Set myXobjBook = PfncobjGetFile(myXstrFullName)
            myXstrBookName = myXobjBook.Name
        '//�G�N�Z���u�b�N�����
            Call PfixCloseExcelBook(myXbisExitFlag, myXstrBookName)
            If myXbisExitFlag = True Then Exit Sub
        '//�t�@�C���p�X���w�肵�ăG�N�Z���u�b�N���J��
            Set myXobjBook = PfncobjOpenExcelBook(myXstrFullName)
        Case 3
        '//�w��t�@�C���̃I�u�W�F�N�g���擾
            Set myXobjBook = PfncobjGetFile(myXstrFullName)
            myXstrBookName = myXobjBook.Name
        '//�w�薼�̃G�N�Z���u�b�N�����ɊJ���Ă���΃u�b�N�I�u�W�F�N�g���擾
            Set myXobjBook = PfncobjGetExcelBookIfOpened(myXstrBookName)
        Case Else
            Exit Sub
    End Select
End Sub

 '��^�e_�w��G�N�Z���u�b�N�����ɊJ���Ă��邩�m�F����
Private Function PfnclonCheckExcelBookOpening( _
            ByVal myXstrFullName As String) As Long
'PfnclonCheckExcelBookOpening = 0 : �w��u�b�N�����݂��Ȃ�
'PfnclonCheckExcelBookOpening = 1 : �J���Ă��Ȃ�
'PfnclonCheckExcelBookOpening = 2 : �w��u�b�N�Ɠ��ꖼ�̕ʃu�b�N���J���Ă���
'PfnclonCheckExcelBookOpening = 3 : �w��u�b�N���J���Ă���
    PfnclonCheckExcelBookOpening = Empty
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    If myXobjFSO.FileExists(myXstrFullName) = False Then Exit Function
  Dim myXstrBookName As String
    myXstrBookName = myXobjFSO.GetFileName(myXstrFullName)
    On Error GoTo ExitPath
  Dim myXstrTmp As String: myXstrTmp = Workbooks(myXstrBookName).FullName
    On Error GoTo 0
    If myXstrTmp = myXstrFullName Then
        PfnclonCheckExcelBookOpening = 3
    Else
        PfnclonCheckExcelBookOpening = 2
    End If
    Set myXobjFSO = Nothing
    Exit Function
ExitPath:
    PfnclonCheckExcelBookOpening = 1
    Set myXobjFSO = Nothing
End Function

 '��^�e_�t�@�C���p�X���w�肵�ăG�N�Z���u�b�N���J��
Private Function PfncobjOpenExcelBook( _
            ByVal myXstrFullName As String) As Object
    Set PfncobjOpenExcelBook = Nothing
    On Error Resume Next
    Set PfncobjOpenExcelBook = Workbooks.Open(myXstrFullName)
    On Error GoTo 0
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

 '��^�e_�w�薼�̃G�N�Z���u�b�N�����ɊJ���Ă���΃u�b�N�I�u�W�F�N�g���擾����
Private Function PfncobjGetExcelBookIfOpened( _
            ByVal myXstrBookName As String) As Object
    Set PfncobjGetExcelBookIfOpened = Nothing
    On Error GoTo ExitPath
    Set PfncobjGetExcelBookIfOpened = Workbooks(myXstrBookName)
    On Error GoTo 0
ExitPath:
End Function

 '��^�o_�G�N�Z���u�b�N�̃v���p�e�B�̃n�C�p�[�����N�̊�_���擾����
Private Function PfncstrGetHyperLinkBase(ByVal myXobjBook As Object) As String
    PfncstrGetHyperLinkBase = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkBase = myXobjBook.BuiltinDocumentProperties("Hyperlink base").Value
    On Error GoTo 0
ExitPath:
End Function

 '��^�o_�G�N�Z���u�b�N�̃v���p�e�B�̃n�C�p�[�����N�̊�_��ݒ肷��
Private Sub PfixSetHyperLinkBase(myXbisExitFlag As Boolean, _
            ByVal myXobjBook As Object, ByVal myXstrHypLnkBase As String)
  Const coXstrBkPrptyName As String = "Hyperlink base"
    myXbisExitFlag = False
    If myXstrHypLnkBase = "" Then Exit Sub
    On Error GoTo ExitPath
    myXobjBook.BuiltinDocumentProperties(coXstrBkPrptyName).Value = myXstrHypLnkBase
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '��^�e_�w��Z���͈͂ɐݒ肳�ꂽ�n�C�p�[�����N��̃p�X���擾����
Private Function PfncstrGetHyperLinkPathAtRange(ByVal myXobjRange As Object) As String
    PfncstrGetHyperLinkPathAtRange = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkPathAtRange = myXobjRange.Hyperlinks(1).Address
    On Error GoTo 0
ExitPath:
End Function

 '��^�e_���΃t�@�C���p�X���w�肵�Đ�΃p�X���擾����
Private Function PfncstrGetAbsolutePath( _
            ByVal myXstrRltvPath As String, ByVal myXobjBook As Object) As String
    PfncstrGetAbsolutePath = Empty
    If myXstrRltvPath = "" Then Exit Function
    If myXobjBook Is Nothing Then Exit Function
  Dim myXstrAbsltPath As String
  Dim myXstrPrntPath As String, myXstrChldPath As String
    myXstrPrntPath = myXobjBook.Path
    myXstrChldPath = myXstrRltvPath
    myXstrPrntPath = Replace(myXstrPrntPath, "/", "\")
    myXstrChldPath = Replace(myXstrChldPath, "/", "\")
  Dim i As Long, j As Long, m As Long, n As Long: m = 0: n = 0
    If Left(myXstrChldPath, Len("..")) = ".." Then
        For i = 1 To Len(myXstrPrntPath)
            If Mid(myXstrPrntPath, i, Len("\")) = "\" Then m = m + 1
        Next i
        For j = 1 To Len(myXstrChldPath)
            If Mid(myXstrChldPath, i, Len("..")) = ".." Then n = n + 1
        Next j
        If m >= n Then
          Dim myXobjFSO As Object
            Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
            Do While Left(myXstrChldPath, Len("..")) = ".."
                myXstrPrntPath = myXobjFSO.GetParentFolderName(myXstrPrntPath)
                myXstrChldPath = Mid(myXstrChldPath, Len("..") + 2)
            Loop
            Set myXobjFSO = Nothing
        Else
            Exit Function
        End If
    End If
    Select Case myXstrChldPath
        Case "": myXstrAbsltPath = myXstrPrntPath
        Case Else: myXstrAbsltPath = myXstrPrntPath & "\" & myXstrChldPath
    End Select
'    Debug.Print "�e�p�X: " & myXstrPrntPath
'    Debug.Print "�q�p�X: " & myXstrChldPath
'    Debug.Print "��΃p�X: " & myXstrAbsltPath
    PfncstrGetAbsolutePath = myXstrAbsltPath
End Function

 '��^�e_�w��t�H���_�̑��݂��m�F����
Private Function PfncbisCheckFolderExist(ByVal myXstrDirPath As String) As Boolean
    PfncbisCheckFolderExist = False
    If myXstrDirPath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFolderExist = myXobjFSO.FolderExists(myXstrDirPath)
    Set myXobjFSO = Nothing
End Function
 
 '��^�e_�w��t�@�C���̑��݂��m�F����
Private Function PfncbisCheckFileExist(ByVal myXstrFilePath As String) As Boolean
    PfncbisCheckFileExist = False
    If myXstrFilePath = "" Then Exit Function
  Dim myXobjFSO As Object: Set myXobjFSO = CreateObject("Scripting.FileSystemObject")
    PfncbisCheckFileExist = myXobjFSO.FileExists(myXstrFilePath)
    Set myXobjFSO = Nothing
End Function

 '��^�o_�w��Z���͈͂Ƀn�C�p�[�����N��ݒ肷��
Private Sub PfixSetHyperLinkWithSheetCellAtRange(myXbisExitFlag As Boolean, _
            ByVal myXobjRange As Object, ByVal myXstrHypLnkAdrs As String, _
            ByVal myXstrSubAdrs As String, ByVal myXstrTxt As String)
'myXstrSubAdrs : "�V�[�g��!�Z���ʒu"
    myXbisExitFlag = False
    If myXobjRange Is Nothing Then Exit Sub
    If myXstrHypLnkAdrs = "" And myXstrSubAdrs = "" Then Exit Sub
    If myXstrTxt = "" Then myXstrTxt = myXobjRange.Value
    On Error GoTo ExitPath
    Call myXobjRange.Worksheet.Hyperlinks.Add( _
            Anchor:=myXobjRange, Address:=myXstrHypLnkAdrs, _
            SubAddress:=myXstrSubAdrs, TextToDisplay:=myXstrTxt)
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '��^�e_�w��}�`�ɐݒ肳�ꂽ�n�C�p�[�����N��̃p�X���擾����
Private Function PfncstrGetHyperLinkPathAtShape(ByVal myXobjShape As Object) As String
    PfncstrGetHyperLinkPathAtShape = Empty
    On Error GoTo ExitPath
    PfncstrGetHyperLinkPathAtShape = myXobjShape.Hyperlink.Address
    On Error GoTo 0
ExitPath:
End Function

 '��^�o_�w��}�`�Ƀn�C�p�[�����N��ݒ肷��
Private Sub PfixSetHyperLinkAtShape(myXbisExitFlag As Boolean, _
            ByVal myXobjShape As Object, ByVal myXstrHypLnkAdrs As String)
    myXbisExitFlag = False
    If myXstrHypLnkAdrs = "" Then Exit Sub
    On Error GoTo ExitPath
    Call myXobjShape.Parent.Hyperlinks _
        .Add(Anchor:=myXobjShape, Address:=myXstrHypLnkAdrs)
    On Error GoTo 0
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

 '��^�o_�G�N�Z���u�b�N���㏑���ۑ�����
Private Sub PfixOverwriteSaveExcelBook(myXbisExitFlag As Boolean, _
            ByVal myXstrBookName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Workbooks(myXstrBookName).Save
    On Error GoTo 0
    GoTo ExitPath
ErrPath:
    myXbisExitFlag = True
ExitPath:
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
End Sub

 '��^�o_�G�N�Z���u�b�N�����
Private Sub PfixCloseExcelBook(myXbisExitFlag As Boolean, _
            ByVal myXstrBookName As String)
    If Application.DisplayAlerts = True Then Application.DisplayAlerts = False
    myXbisExitFlag = False
    On Error GoTo ErrPath
    Workbooks(myXstrBookName).Close SaveChanges:=False
    On Error GoTo 0
    GoTo ExitPath
ErrPath:
    myXbisExitFlag = True
ExitPath:
    If Application.DisplayAlerts = False Then Application.DisplayAlerts = True
End Sub

'Dummy�o_
Private Sub MsubDummy()
End Sub

'===============================================================================================

'��ModuleProc��_�G�N�Z���u�b�N���̑S�V�[�g���S�Z���͈́��S�}�`�ɑ΂��ď��������s����
Private Sub callMexeHypLnkToAbslt()
  Dim myXbisCompFlag As Boolean
    Call MexeHypLnkToAbslt.callProc(myXbisCompFlag)
    Debug.Print "����: " & myXbisCompFlag
End Sub
