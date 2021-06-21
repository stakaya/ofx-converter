Attribute VB_Name = "CUIMain"
Option Explicit

Sub Main()

   '２重起動の判定
    If App.PrevInstance Then End
    
    Dim Table()     As String
    Dim InputFile   As String
    Dim OutputFile  As String
    Dim IsOverWrite As Boolean
    Dim Temp        As Variant
    Dim Ofx         As Object
    
    Set Ofx = CreateObject("OfxCore.Ofx")
    
    'コマンドライン引数取り出し
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
    
    '必須チェック
    If InputFile = Empty Then
        Dim Message As String
        Message = "コマンドラインオプションに誤りがあります。" & vbCrLf & _
                  "入力ファイルは必須です。"
        MsgBox Message, vbExclamation, "OFXツール"
        End
    End If

    'ファイル生成
    Call Ofx.Convert(InputFile, OutputFile, IsOverWrite)
End Sub

