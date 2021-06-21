VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ofx"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"47B0A2170330"
'
Option Base 0

Option Explicit

' ����� 0 �ɐݒ肷��ƁA���̸׽�����ޯ�޺��ނ������ɂȂ�܂�
#Const DebugMode = 1
#If DebugMode Then
    '�Ő������ꂽ�A������׽ID��ێ�����Ǐ��ϐ� Class_Initialize
    '##ModelId=47B0A2170362
    Private mlClassDebugID As Long
#End If

'##ModelId=47B0A2170365
Private Log As New Logger
'##ModelId=47B0A217036E
Private Ini As New IniFile

Private Response As String
Private xDoc As Object

'##ModelId=47B0A217036F
Public Function Convert(InputFile As String, Optional OutputFile As String = "", Optional IsUpdate As Boolean = False) As Boolean
    On Error GoTo EncodeErr

    InputFile = Replace(InputFile, """", Empty)
    OutputFile = Replace(OutputFile, """", Empty)

    '���̓t�@�C���`�F�b�N
    If InputFile = Empty Or Dir(InputFile) = Empty Then
        Log.Error "���̓t�@�C��(" & InputFile & ")�����݂��܂���B"
        Convert = False
        Exit Function
    End If

    Set xDoc = CreateObject("MSXML.DOMDocument")
    If Not xDoc.Load(InputFile) Then
        ' �h�L�������g�ɓǂݍ��݂Ɏ��s���܂����B
        Dim xPE As Object
        Set xPE = xDoc.parseError
        With xPE
        Log.Error "���� XML �h�L�������g�̓ǂݍ��݂Ɏ��s���܂��� :" & _
            "���̃G���[�������ł� :" & vbLf & _
            "�G���[ # : " & .errorCode & ": " & xPE.reason & _
            "�s #: " & .Line & vbLf & _
            "�s�ʒu : " & .linepos & vbLf & _
            "�t�@�C�����̈ʒu : " & .filepos & vbLf & _
            "�\�[�X �e�L�X�g : " & .srcText & vbLf & _
            "�h�L�������g URL : " & .url
        End With
        
        Convert = False
    End If

    ' Write OFX tags
    WriteOFX
    
    '�t�@�C���Ɏw�肪�����ꍇ
    If OutputFile = Empty Then
        OutputFile = Format(Now, "yyyymmddhhMMss") & ".ofx"
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .WriteText Response
            .SaveToFile OutputFile, 2
            .Close
        End With
    ElseIf IsUpdate Or Dir(OutputFile) = Empty Then
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .WriteText Response
            .SaveToFile OutputFile, 2
            .Close
        End With
    Else
        Log.Info OutputFile & "�����݂����ׁA�t�@�C���͍X�V����܂���ł����B"
        Convert = False
        Exit Function
    End If

    ' ���L�̏�����ADODB.Stream�ɂăS�~������̂�����
    Dim Data()   As Byte
    Dim Out()    As Byte
    Dim FileNo   As Integer
    Dim FileSize As Long
    Dim Written  As Long

    FileSize = FileLen(OutputFile) - 1
    ReDim Data(FileSize)
    FileNo = FreeFile

    '�t�@�C�����o�C�i�����[�h�ŊJ��
    Open OutputFile For Binary As #FileNo
    '�o�C�i���f�[�^�Ƃ��ēǍ����
    Get #FileNo, , Data
    Close #FileNo

    Dim i As Integer
    For i = 0 To FileSize
        If Asc("O") = Data(i) Then
            Exit For
        End If
    Next i
    
    ReDim Out(FileSize - i)
    Dim j As Integer
    For j = 0 To FileSize - i
       Out(j) = Data(i + j)
    Next j

    '�t�@�C���폜
    Kill OutputFile
    FileNo = FreeFile
    Open OutputFile For Binary Access Write As #FileNo
    Put #FileNo, , Out
    Close #FileNo

    Convert = True
    Exit Function
EncodeErr:
    Log.Error "�t�@�C���̕ϊ��Ɏ��s���܂����B" & Err.Description
    
    Call RaiseError(MyUnhandledError, "Ofx:Convert Method")
End Function

'##ModelId=47B0A2170373
Private Sub Class_Initialize()

    #If DebugMode Then
        '���̗L���ȸ׽ID���擾���A���̸׽�����܂��������ꂽ���Ƃ��o�͂��܂��B
        '
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=47B0A2170376
Private Sub Class_Terminate()

    
    #If DebugMode Then
    '�׽�͔j������܂��B
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=47B0A2170377
    Public Property Get ClassDebugID() As Long
        '���ޯ��Ӱ�ނł���΁A�ݼ��ς��⍇���ł��邱�������è��\�ɏo���܂�
        ClassDebugID = mlClassDebugID
    End Property
#End If

Private Sub WriteUTF8(s)
    Response = Response & s
End Sub

Private Sub BlockBegin(s)
    Response = Response & "<" & s & ">" & vbLf
End Sub

Private Sub BlockEnd(s)
    Response = Response & "</" & s & ">" & vbLf
End Sub

Private Sub WriteTag(tag, strValue)
    Dim cchMax As Integer

    Select Case tag
    Case "ORG"
        cchMax = 32
    Case "BANKID"
        cchMax = 9
    Case "BRANCHID"
        cchMax = 22
    Case "ACCTID"
        cchMax = 22
    Case "FITID"
        cchMax = 255
    Case "NAME"
        cchMax = 32
    Case "MEMO"
        cchMax = 255
    Case Else
        cchMax = 0
    End Select

    ' Trim string to appropriate length if specified
    strValue = Trim(strValue)
    If cchMax > 0 Then strValue = Trim(Left(strValue, cchMax))

    ' Write in UTF-8
    WriteUTF8 "<" & tag & ">" & strValue & vbLf
End Sub

' --------------------------------
' Format date in OFX format in JST
Private Function StrFormatOFXDate(d)
    StrFormatOFXDate = Year(d)
    StrFormatOFXDate = StrFormatOFXDate & Right("00" & Month(d), 2)
    StrFormatOFXDate = StrFormatOFXDate & Right("00" & Day(d), 2)
    StrFormatOFXDate = StrFormatOFXDate & Right("00" & Hour(d), 2)
    StrFormatOFXDate = StrFormatOFXDate & Right("00" & Minute(d), 2)
    StrFormatOFXDate = StrFormatOFXDate & Right("00" & Second(d), 2)
    StrFormatOFXDate = StrFormatOFXDate & "[+9:JST]"
End Function

Private Sub WriteOFX()
    Dim node As Object
    Dim nodes As Object

    ' Write OFX header
    Response = "OFXHEADER:100" & vbLf & _
        "DATA:OFXSGML" & vbLf & _
        "VERSION:102" & vbLf & _
        "SECURITY:NONE" & vbLf & _
        "ENCODING:UTF-8" & vbLf & _
        "CHARSET:CSUNICODE" & vbLf & _
        "COMPRESSION:NONE" & vbLf & _
        "OLDFILEUID:NONE" & vbLf & _
        "NEWFILEUID:NONE" & vbLf & vbLf
   
    BlockBegin "OFX"
    WriteSignOnMsgsRs "�}�l�[�R���o�[�^"
    
    Set nodes = xDoc.getElementsByTagName("BANKMSGSRSV1")
    If nodes.length > 0 Then
        BlockBegin "BANKMSGSRSV1"
        For Each node In xDoc.selectNodes("/OFX/BANKMSGSRSV1/STMTTRNRS")
            WriteStmtTrnRs node
        Next node
        BlockEnd "BANKMSGSRSV1"
    End If
    
    Set nodes = xDoc.getElementsByTagName("CREDITCARDMSGSRSV1")
    If nodes.length > 0 Then
        BlockBegin "CREDITCARDMSGSRSV1"
        For Each node In xDoc.selectNodes("/OFX/CREDITCARDMSGSRSV1/CCSTMTTRNRS")
            WriteCCStmtTrnRs node
        Next node
        BlockEnd "CREDITCARDMSGSRSV1"
    End If
    
    BlockEnd "OFX"
End Sub

Private Sub WriteSignOnMsgsRs(strFiName)
    BlockBegin "SIGNONMSGSRSV1"
    BlockBegin "SONRS"
    WriteStatusOK
    WriteTag "DTSERVER", StrFormatOFXDate(Now)
    WriteTag "LANGUAGE", "JPN"
    If strFiName <> "" Then
        BlockBegin "FI"
        WriteTag "ORG", strFiName
        BlockEnd "FI"
    End If
    BlockEnd "SONRS"
    BlockEnd "SIGNONMSGSRSV1"
End Sub

Private Sub WriteStatusOK()
    Response = Response & "<STATUS><CODE>0<SEVERITY>INFO</STATUS>" & vbLf
End Sub

' --------------------------------------
' STMTTRNRS -- bank statement main block
Private Sub WriteStmtTrnRs(node As Object)
    BlockBegin "STMTTRNRS"
    WriteTag "TRNUID", 0
    WriteStatusOK
    WriteStmtRs "STMTRS", node
    BlockEnd "STMTTRNRS"
End Sub

' -----------------------------------------------
' CCSTMTTRNRS -- credit card statement main block
Private Sub WriteCCStmtTrnRs(node As Object)
    BlockBegin "CCSTMTTRNRS"
    WriteTag "TRNUID", 0
    WriteStatusOK
    WriteStmtRs "CCSTMTRS", node
    BlockEnd "CCSTMTTRNRS"
End Sub

' ----------------------------------------------
' STMTRS/CCSTMTRS -- bank/credit card statements
Private Sub WriteStmtRs(sSTMTRS, node As Object)
    BlockBegin sSTMTRS
    WriteTag "CURDEF", "JPY"
    WriteAcctFrom node
    
    Dim doc As Object
    Dim temp As Object
    Dim work As Object
    Dim stmtt As Object
    Set doc = CreateObject("MSXML.DOMDocument")
    doc.loadXML (node.xml)
    
    For Each temp In doc.getElementsByTagName("BANKTRANLIST")
        BlockBegin "BANKTRANLIST"
        For Each work In temp.childNodes
            If work.nodeName = "STMTTRN" Then
                BlockBegin "STMTTRN"
                For Each stmtt In work.childNodes
                    WriteTag stmtt.nodeName, stmtt.Text
                Next stmtt
                BlockEnd "STMTTRN"
            Else
                WriteTag work.nodeName, work.Text
            End If
        Next work
        BlockEnd "BANKTRANLIST"
    Next temp

    WriteBalance node
    BlockEnd sSTMTRS
End Sub

Private Sub WriteAcctFrom(node As Object)

    Dim doc As Object
    Dim temp As Object
    Dim work As Object
    Set doc = CreateObject("MSXML.DOMDocument")
    doc.loadXML (node.xml)
    
    For Each temp In doc.getElementsByTagName("BANKACCTFROM")
        BlockBegin "BANKACCTFROM"
        For Each work In temp.childNodes
            WriteTag work.nodeName, work.Text
        Next work
        BlockEnd "BANKACCTFROM"
    Next temp
    
    For Each temp In doc.getElementsByTagName("CCACCTFROM")
        BlockBegin "CCACCTFROM"
        For Each work In temp.childNodes
            WriteTag work.nodeName, work.Text
        Next work
        BlockEnd "CCACCTFROM"
    Next temp
End Sub


' -------------------------------------
' LEDGARBAL -- bank/credit card balance
Private Sub WriteBalance(node As Object)
    Dim doc As Object
    Dim temp As Object
    Set doc = CreateObject("MSXML.DOMDocument")
    doc.loadXML (node.xml)
    
    For Each temp In doc.getElementsByTagName("LEDGERBAL")
        BlockBegin "LEDGERBAL"
        WriteTag temp.firstChild.nodeName, temp.firstChild.Text
        WriteTag temp.lastChild.nodeName, temp.lastChild.Text
        BlockEnd "LEDGERBAL"
    Next temp
    
    For Each temp In doc.getElementsByTagName("AVAILBAL")
        BlockBegin "AVAILBAL"
        WriteTag temp.firstChild.nodeName, temp.firstChild.Text
        WriteTag temp.lastChild.nodeName, temp.lastChild.Text
        BlockEnd "AVAILBAL"
    Next temp
End Sub
