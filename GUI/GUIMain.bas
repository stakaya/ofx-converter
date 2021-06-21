Attribute VB_Name = "GUIMain"
Option Explicit

Public Const OFN_READONLY = &H1                '�ǂݎ���pON
Public Const OFN_OVERWRITEPROMPT = &H2         '�㏑���m�F
Public Const OFN_HIDEREADONLY = &H4            '�ǂݎ���pOFF
Public Const OFN_SHOWHELP = &H10               '�������ݕ\��
Public Const OFN_ALLOWMULTISELECT = &H200      '�����I����
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800         '���݂��Ȃ��߽����͕s��
Public Const OFN_FILEMUSTEXIST = &H1000        '���݂��Ȃ�̧�ق���͕s��
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_EXPLORER = &H80000

Public Type OPENFILENAME
    lStructSize       As Long   '�\���̻̂���
    hWndOwner         As Long   '�޲�۸��ޯ���̐e����޳�������
    hInstance         As Long   'APP�ݽ�ݽ
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String '�I�����ꂽ̧�ٖ�
    nMaxFile          As Long   '̧�ٖ��̍ő咷
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
(pOpenfilename As OPENFILENAME) As Long

Public Ofx As Object

Sub Main()

   '�Q�d�N���̔���
    If App.PrevInstance Then End
    Dim FileName As OPENFILENAME
    FileName.lStructSize = Len(FileName)
    FileName.hWndOwner = Money.hWnd
    FileName.hInstance = App.hInstance
    FileName.lpstrFilter = "*.ofx" & vbNullChar
    FileName.nFilterIndex = 1
    FileName.lpstrFile = String(256, Chr(0))
    FileName.nMaxFile = 256
    FileName.lpstrFileTitle = String(256, Chr(0))
    FileName.nMaxFileTitle = 256
    FileName.lpstrInitialDir = CurDir
    FileName.lpstrTitle = "̧�ق��J��"
    FileName.flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    
    If GetOpenFileName(FileName) <> 0 Then
        Set Ofx = CreateObject("OfxCore.Ofx")
        
        Dim InputFile As String
        InputFile = Left$(FileName.lpstrFile, _
                    InStr(FileName.lpstrFile, vbNullChar) - 1)

        '�t�@�C������
        Call Ofx.Convert(InputFile)
    End If

    '�I��
    End
End Sub



