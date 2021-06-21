Attribute VB_Name = "CUIMain"
Option Explicit

Sub Main()

   '�Q�d�N���̔���
    If App.PrevInstance Then End
    
    Dim Table()     As String
    Dim InputFile   As String
    Dim OutputFile  As String
    Dim IsOverWrite As Boolean
    Dim Temp        As Variant
    Dim Ofx         As Object
    
    Set Ofx = CreateObject("OfxCore.Ofx")
    
    '�R�}���h���C���������o��
    Table = Split(Command, "-")
    For Each Temp In Table
         Select Case UCase(Left(Temp, 1))
             Case "I"
                 InputFile = Trim(Mid(Temp, 2, Len(Temp)))
             Case "O"
                 OutputFile = Trim(Mid(Temp, 2, Len(Temp)))
             Case "U"
                 IsOverWrite = True
         End Select
    Next Temp
    
    '�K�{�`�F�b�N
    If InputFile = Empty Then
        Dim Message As String
        Message = "�R�}���h���C���I�v�V�����Ɍ�肪����܂��B" & vbCrLf & _
                  "���̓t�@�C���͕K�{�ł��B"
        MsgBox Message, vbExclamation, "OFX�c�[��"
        End
    End If

    '�t�@�C������
    Call Ofx.Convert(InputFile, OutputFile, IsOverWrite)
End Sub

