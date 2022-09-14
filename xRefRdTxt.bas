Attribute VB_Name = "xRefRdTxt"
'Includes CRdTxtNoOpn
'Includes CRdTxtNoOpnUTF8
'Includes CRdTxtOpn
'Includes PfncstrGetTextFileCharset
'Includes PfixChangeModuleConstValue
'Includes x

Option Explicit
Option Base 1

'��ModuleProc��_�e�L�X�g�t�@�C���̓��e���擾����
'Rev.001
  
'//���W���[��������
  Private Const meMstrMdlName As String = "xRefRdTxt"
  Private Const meMlonExeNum As Long = 0
  
'//���W���[�����萔
  Private Const coXstrANSI As Variant = "Shift_JIS (ANSI)"
  Private Const coXstrUTF8 As Variant = "UTF-8"
  
'//�o�͐���M��
  Private myXbisCmpltFlag As Boolean
  
'//�o�̓f�[�^
  Private myXstrFileCharset As String
    'myXstrFileCharset = Shift_JIS (ANSI)
    'myXstrFileCharset = UTF-8
    'myXstrFileCharset = UTF-8 BOM
    'myXstrFileCharset = UTF-16 LE BOM
    'myXstrFileCharset = UTF-16 BE BOM
    'myXstrFileCharset = EUC-JP
  Private myXstrDirPath As String, myXstrFileName As String, _
            myXstrBaseName As String, myXstrExtsn As String
  Private myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
            myZstrTxtData() As String
    'myZstrTxtData(i, j) : �e�L�X�g�t�@�C�����e
  
'//���̓f�[�^
  Private myXstrOrgFilePath As String
  Private myXlonBgn As Long, myXlonEnd As Long, _
            myXbisSpltOptn As Boolean, myXstrInSpltChr As String
  
'//���W���[�����ϐ�_����M��
  Private myXbisExitFlag As Boolean
  
'//���W���[�����ϐ�_�f�[�^

'iniP_���W���[�����ϐ�������������
Private Sub initializeModuleVariables()
    myXbisExitFlag = False
    
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
    Call callxRefRdTxt
    
'//�������ʕ\��
    Select Case myXbisCmpltFlag
        Case True: MsgBox "���s����"
        Case Else: MsgBox "�ُ킠��", vbExclamation
    End Select
    
End Sub

'PublicP_
Public Sub callProc( _
            myXbisCmpltFlagOUT As Boolean, _
            myXstrDirPathOUT As String, myXstrFileNameOUT As String, _
            myXstrBaseNameOUT As String, myXstrExtsnOUT As String, _
            myXlonTxtRowCntOUT As Long, myXlonTxtColCntOUT As Long, _
            myZstrTxtDataOUT() As String, _
            ByVal myXstrOrgFilePathIN As String, _
            ByVal myXlonBgnIN As Long, ByVal myXlonEndIN As Long, _
            ByVal myXbisSpltOptnIN As Boolean, ByVal myXstrInSpltChrIN As String)

'//���͕ϐ���������
    myXstrOrgFilePath = Empty
    myXlonBgn = Empty: myXlonEnd = Empty
    myXbisSpltOptn = False: myXstrInSpltChr = Empty

'//���͕ϐ�����荞��
    myXstrOrgFilePath = myXstrOrgFilePathIN
    myXlonBgn = myXlonBgnIN
    myXlonEnd = myXlonEndIN
    myXbisSpltOptn = myXbisSpltOptnIN
    myXstrInSpltChr = myXstrInSpltChrIN
    
'//�o�͕ϐ���������
    myXbisCmpltFlagOUT = False
    
    myXstrDirPathOUT = Empty: myXstrFileNameOUT = Empty
    myXstrBaseNameOUT = Empty: myXstrExtsnOUT = Empty
    myXlonTxtRowCntOUT = Empty: myXlonTxtColCntOUT = Empty
    Erase myZstrTxtDataOUT

'//�������s
    Call ctrProc
    If myXbisCmpltFlag = False Then Exit Sub

'//�o�͕ϐ��Ɋi�[
    myXbisCmpltFlagOUT = myXbisCmpltFlag
    
    myXstrDirPathOUT = myXstrDirPath
    myXstrFileNameOUT = myXstrFileName
    myXstrBaseNameOUT = myXstrBaseName
    myXstrExtsnOUT = myXstrExtsn
    myXlonTxtRowCntOUT = myXlonTxtRowCnt
    myXlonTxtColCntOUT = myXlonTxtColCnt
    If myXlonTxtRowCntOUT <= 0 Or myXlonTxtColCntOUT <= 0 Then Exit Sub
    myZstrTxtDataOUT() = myZstrTxtData()

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
    
'//S:�e�L�X�g�t�@�C���̓��e���擾
    Call snsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "3"     'PassFlag
    
'//P:
    Call prsProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "4"     'PassFlag
    
'//Run:
    Call runProc
    If myXbisExitFlag = True Then GoTo ExitPath
'    Debug.Print "PassFlag: " & meMstrMdlName & "5"     'PassFlag
    
    Call checkOutputVariables: If myXbisExitFlag = True Then GoTo ExitPath
    myXbisCmpltFlag = True
ExitPath:
    If coXbisTestMode = False Then Call recProc
    Call initializeModuleVariables
End Sub

'iniP_�o�͕ϐ�������������
Private Sub initializeOutputVariables()
    myXbisCmpltFlag = False
    myXstrFileCharset = Empty
    myXstrDirPath = Empty: myXstrFileName = Empty
    myXstrBaseName = Empty: myXstrExtsn = Empty
    myXlonTxtRowCnt = Empty: myXlonTxtColCnt = Empty: Erase myZstrTxtData
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
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'SetP_����p�ϐ���ݒ肷��
Private Sub setControlVariables()
  
  Dim myXstrPrntPath As String, myXstrFileName As String
    myXstrPrntPath = ThisWorkbook.Path
    myXstrFileName = "testIN.txt"
    myXstrOrgFilePath = myXstrPrntPath & "\" & myXstrFileName
    
    myXlonBgn = 1
    myXlonEnd = 0
    
    myXbisSpltOptn = True
    myXstrInSpltChr = ""
    'myXbisSpltOptn = True  : ������𕪊���������
    'myXbisSpltOptn = False : ������𕪊��������Ȃ�
    
End Sub

'SnsP_
Private Sub snsProc()
    myXbisExitFlag = False
    
'//�w��e�L�X�g�t�@�C���̕����R�[�h���擾
    myXstrFileCharset = PfncstrGetTextFileCharset(myXstrOrgFilePath)
    If myXstrFileCharset = "" Then GoTo ExitPath
    
'//�����R�[�h�ŏ����𕪊�
    Select Case myXstrFileCharset
        Case coXstrANSI
        '//�t�@�C�����J�����Ƀe�L�X�g�t�@�C���̓��e���擾
            Call instCRdTxtNoOpn
            
        Case coXstrUTF8
        '//�t�@�C�����J������UTF8�`���e�L�X�g�t�@�C���̓��e���擾
            Call instCRdTxtNoOpnUTF8
            
        Case Else
        '//�t�@�C�����J���ăe�L�X�g�t�@�C���̓��e���擾
            Call instCRdTxtOpn
            
    End Select
    If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo ExitPath
    
    Exit Sub
ExitPath:
    myXbisExitFlag = True
End Sub

'PrcsP_
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

'��ClassProc��_�t�@�C�����J�����Ƀe�L�X�g�t�@�C���̓��e���擾����
Private Sub instCRdTxtNoOpn()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtNoOpn As CRdTxtNoOpn: Set myXinsRdTxtNoOpn = New CRdTxtNoOpn
    With myXinsRdTxtNoOpn
    '//�N���X���ϐ��ւ̓���
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtNoOpn = Nothing
End Sub

'��ClassProc��_�t�@�C�����J������UTF8�`���e�L�X�g�t�@�C���̓��e���擾����
Private Sub instCRdTxtNoOpnUTF8()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtNoOpn As CRdTxtNoOpnUTF8: Set myXinsRdTxtNoOpn = New CRdTxtNoOpnUTF8
    With myXinsRdTxtNoOpn
    '//�N���X���ϐ��ւ̓���
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtNoOpn = Nothing
End Sub

'��ClassProc��_�t�@�C�����J���ăe�L�X�g�t�@�C���̓��e���擾����
Private Sub instCRdTxtOpn()
  Dim Lc As Long
  Dim myZlonTmp(1) As Long, Lo As Long: Lo = LBound(myZlonTmp)
  Dim i As Long, j As Long
  Dim myXinsRdTxtOpn As CRdTxtOpn: Set myXinsRdTxtOpn = New CRdTxtOpn
    With myXinsRdTxtOpn
    '//�N���X���ϐ��ւ̓���
    '//�e�L�X�g�t�@�C���p�X���w��
        .letFilePath = myXstrOrgFilePath
        .letRdBgnEnd(1) = myXlonBgn
        .letRdBgnEnd(2) = myXlonEnd
        .letSpltOptn = myXbisSpltOptn
        .letSpltChr = myXstrInSpltChr
    '//�N���X���v���V�[�W���̎��s�ƃN���X���ϐ�����̏o��
        .exeProc
        myXstrDirPath = .getDirPath
        myXstrFileName = .getFileName
        myXstrBaseName = .getBaseName
        myXstrExtsn = .getExtsn
        myXlonTxtRowCnt = .getTxtRowCnt
        myXlonTxtColCnt = .getTxtColCnt
        If myXlonTxtRowCnt <= 0 Or myXlonTxtColCnt <= 0 Then GoTo JumpPath
        i = myXlonTxtRowCnt + Lo - 1: j = myXlonTxtColCnt + Lo - 1
        ReDim myZstrTxtData(i, j) As String
        Lc = .getOptnBase
        For i = 1 To myXlonTxtRowCnt
            For j = 1 To myXlonTxtColCnt
                myZstrTxtData(i + Lo - 1, j + Lo - 1) = .getTxtDataAry(i + Lc - 1, j + Lc - 1)
            Next j
        Next i
    End With
JumpPath:
    Set myXinsRdTxtOpn = Nothing
End Sub

'===============================================================================================
 
 '��^�e_�w��e�L�X�g�t�@�C���̕����R�[�h���擾����
Private Function PfncstrGetTextFileCharset(ByVal myXstrFilePath As String) As String
'myXstrCharset = Shift_JIS (ANSI)
'myXstrCharset = UTF-8
'myXstrCharset = UTF-8 BOM
'myXstrCharset = UTF-16 LE BOM
'myXstrCharset = UTF-16 BE BOM
'myXstrCharset = EUC-JP
    PfncstrGetTextFileCharset = Empty
  Dim myXstrCharset As String, i As Long
  Dim myXlonHdlFile As Long, myXlonFileLen As Long
  Dim myZbytFile() As Byte, b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
  Dim myXlonSJIS As Long, myXlonUTF8 As Long, myXlonEUC As Long
  Dim myZlonTmp(1) As Long, L As Long: L = LBound(myZlonTmp)
'//�t�@�C���ǂݍ���
    On Error Resume Next
    myXlonFileLen = FileLen(myXstrFilePath)
    ReDim myZbytFile(myXlonFileLen)
    If Err.Number <> 0 Then Exit Function
    myXlonHdlFile = FreeFile()
    Open myXstrFilePath For Binary As #myXlonHdlFile
    Get #myXlonHdlFile, , myZbytFile
    Close #myXlonHdlFile
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
'//BOM�ɂ�锻�f
    If (myZbytFile(L) = &HEF And myZbytFile(L + 1) = &HBB And myZbytFile(L + 2) = &HBF) Then
        myXstrCharset = "UTF-8 BOM"
        GoTo SetPath
    ElseIf (myZbytFile(L) = &HFF And myZbytFile(L + 1) = &HFE) Then
        myXstrCharset = "UTF-16 LE BOM"
        GoTo SetPath
    ElseIf (myZbytFile(L) = &HFE And myZbytFile(L + 1) = &HFF) Then
        myXstrCharset = "UTF-16 BE BOM"
        GoTo SetPath
    End If
'//BINARY
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If (b1 >= &H0 And b1 <= &H8) Or _
                (b1 >= &HA And b1 <= &H9) Or _
                (b1 >= &HB And b1 <= &HC) Or _
                (b1 >= &HE And b1 <= &H19) Or _
                (b1 >= &H1C And b1 <= &H1F) Or _
                (b1 = &H7F) Then
            myXstrCharset = "BINARY"
            GoTo SetPath
        End If
    Next i
'//Shift_JIS
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H9 Or b1 = &HA Or b1 = &HD Or _
                (b1 >= &H20 And b1 <= &H7E) Or _
                (b1 >= &HB0 And b1 <= &HDF) Then
            myXlonSJIS = myXlonSJIS + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If ((b1 >= &H81 And b1 <= &H9F) Or (b1 >= &HE0 And b1 <= &HFC)) And _
                        ((b2 >= &H40 And b2 <= &H7E) Or (b2 >= &H80 And b2 <= &HFC)) Then
                   myXlonSJIS = myXlonSJIS + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
'//UTF-8
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H9 Or b1 = &HA Or b1 = &HD Or (b1 >= &H20 And b1 <= &H7E) Then
            myXlonUTF8 = myXlonUTF8 + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If (b1 >= &HC2 And b1 <= &HDF) And (b2 >= &H80 And b2 <= &HBF) Then
                   myXlonUTF8 = myXlonUTF8 + 2
                   i = i + 1
                Else
                    If (i < myXlonFileLen - 3) Then
                        b3 = myZbytFile(i + 2)
                        If (b1 >= &HE0 And b1 <= &HEF) And _
                                (b2 >= &H80 And b2 <= &HBF) And _
                                (b3 >= &H80 And b3 <= &HBF) Then
                            myXlonUTF8 = myXlonUTF8 + 3
                            i = i + 2
                        Else
                            If (i < myXlonFileLen - 4) Then
                                b4 = myZbytFile(i + 3)
                                If (b1 >= &HF0 And b1 <= &HF7) And _
                                        (b2 >= &H80 And b2 <= &HBF) And _
                                        (b3 >= &H80 And b3 <= &HBF) And _
                                        (b4 >= &H80 And b4 <= &HBF) Then
                                    myXlonUTF8 = myXlonUTF8 + 4
                                    i = i + 3
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
'//EUC-JP
    For i = L To myXlonFileLen + L - 1
        b1 = myZbytFile(i)
        If b1 = &H7 Or b1 = 10 Or b1 = 13 Or (b1 >= &H20 And b1 <= &H7E) Then
            myXlonEUC = myXlonEUC + 1
        Else
            If (i < myXlonFileLen - 2) Then
                b2 = myZbytFile(i + 1)
                If ((b1 >= &HA1 And b1 <= &HFE) And (b2 >= &HA1 And b2 <= &HFE)) Or _
                        (b1 = &H8E And (b2 >= &HA1 And b2 <= &HDF)) Then
                   myXlonEUC = myXlonEUC + 2
                   i = i + 1
                End If
            End If
        End If
    Next i
'//�����R�[�h�o�����ʂɂ�锻�f
    If (myXlonSJIS <= myXlonUTF8) And (myXlonEUC <= myXlonUTF8) Then
        myXstrCharset = "UTF-8"
        GoTo SetPath
    End If
    If (myXlonUTF8 <= myXlonSJIS) And (myXlonEUC <= myXlonSJIS) Then
        myXstrCharset = "Shift_JIS"
        GoTo SetPath
    End If
    If (myXlonUTF8 <= myXlonEUC) And (myXlonSJIS <= myXlonEUC) Then
        myXstrCharset = "EUC-JP"
        GoTo SetPath
    End If
    Exit Function
SetPath:
    PfncstrGetTextFileCharset = myXstrCharset
End Function

 '��^�o_���W���[�����萔�̒l��ύX����
Private Sub PfixChangeModuleConstValue(myXbisExitFlag As Boolean, _
            ByVal myXstrMdlName As String, ByRef myZvarM() As Variant)
    myXbisExitFlag = False
    If myXstrMdlName = "" Then GoTo ExitPath
  Dim L As Long, myXvarTmp As Variant
    On Error GoTo ExitPath
    If IsArray(myZvarM) = False Then GoTo ExitPath
    L = LBound(myZvarM, 1): myXvarTmp = myZvarM(L, L)
    On Error GoTo 0
  Dim myXlonDclrLines As Long
    With ThisWorkbook.VBProject.VBComponents(myXstrMdlName).CodeModule
    myXlonDclrLines = .CountOfDeclarationLines
    If myXlonDclrLines < 1 Then GoTo ExitPath
  Dim i As Long, n As Long
  Dim myXstrTmp As String, myXstrRplcCode As String
    For i = 1 To myXlonDclrLines
        myXstrTmp = .Lines(i, 1)
        For n = LBound(myZvarM, 1) To UBound(myZvarM, 1)
          Dim myXstrSrch As String
            myXstrSrch = "Const" & Space(1) & myZvarM(n, L) & Space(1) & "As" & Space(1)
            If InStr(myXstrTmp, myXstrSrch) = 0 Then GoTo NextPath
          Dim myXstrOrg As String
            myXstrOrg = Left(myXstrTmp, InStr(myXstrTmp, "=" & Space(1)) + 1)
            myXstrRplcCode = myXstrOrg & myZvarM(n, L + 1)
            Application.DisplayAlerts = False
            Call .ReplaceLine(i, myXstrRplcCode)
            Application.DisplayAlerts = True
NextPath:
        Next n
    Next i
    End With
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
'  Dim myXstrPrntPath As String, myXstrFileName As String
'    myXstrPrntPath = ThisWorkbook.Path
'    myXstrFileName = "testIN.txt"
'    myXstrOrgFilePath = myXstrPrntPath & "\" & myXstrFileName
'    myXlonBgn = 1
'    myXlonEnd = 0
'    myXbisSpltOptn = True
'    myXstrInSpltChr = ""
'    'myXbisSpltOptn = True  : ������𕪊���������
'    'myXbisSpltOptn = False : ������𕪊��������Ȃ�
'End Sub
'��ModuleProc��_�e�L�X�g�t�@�C���̓��e���擾����
Private Sub callxRefRdTxt()
'  Dim myXlonOutputOptn As Long, _
'        myXstrOrgFilePath As String, myXlonBgn As Long, myXlonEnd As Long, _
'        myXbisSpltOptn As Boolean, myXstrInSpltChr As String
'  Dim myXbisCmpltFlag As Boolean
'  Dim myXlonTxtRowCnt As Long, myXlonTxtColCnt As Long, _
'        myZstrTxtData() As String
    Call xRefRdTxt.callProc( _
            myXbisCmpltFlag, _
            myXstrDirPath, myXstrFileName, myXstrBaseName, myXstrExtsn, _
            myXlonTxtRowCnt, myXlonTxtColCnt, myZstrTxtData, _
            myXstrOrgFilePath, myXlonBgn, myXlonEnd, myXbisSpltOptn, myXstrInSpltChr)
End Sub
'
'  Public Const coXbisTestMode As Boolean = True
'  Public Const coXbisTestMode As Boolean = False
'
Private Sub resetConstantInxRefRdTxt()
'//xRefRdTxt���W���[���̃��W���[���������̃��Z�b�g����
    Call xRefRdTxt.resetConstant
End Sub
