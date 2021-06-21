VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"47B0A218007E"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
Option Base 0

Option Explicit

'##ModelId=47B0A21800A6
Private Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
 
'##ModelId=47B0A21800B0
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'##ModelId=47B0A21800B7
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
 
'##ModelId=47B0A218009C
Private IniFilePath As String
 
' これを 0 に設定すると、このｸﾗｽのﾃﾞﾊﾞｯｸﾞｺｰﾄﾞが無効になります
#Const DebugMode = 1
#If DebugMode Then
    'で生成された連続するｸﾗｽIDを保持する局所変数 Class_Initialize
    '##ModelId=47B0A218009D
    Private mlClassDebugID As Long
#End If

'##ModelId=47B0A21800BE
Private Sub Class_Terminate()
    #If DebugMode Then
    'ｸﾗｽは破棄されます。
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=47B0A21800BF
Private Sub Class_Initialize()

    'INIファイルのフルパスを取得
    If Right$(App.Path, 1) = "\" Then
        IniFilePath = App.Path & App.EXEName & ".ini"
    Else
        IniFilePath = App.Path & "\" & App.EXEName & ".ini"
    End If

    #If DebugMode Then
        '次の有効なｸﾗｽIDを取得し、このｸﾗｽがうまく生成されたことを出力します。
        '
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=47B0A21800C4
Public Function ReadInteger(ByVal Section As String, ByVal Key As String, ByVal Default As Long) As Long
    On Error GoTo ReadIntegerErr
    
    ReadInteger = GetPrivateProfileInt(Section, Key, Default, IniFilePath)
    
    Exit Function
ReadIntegerErr:
    Call RaiseError(MyUnhandledError, "IniFile:ReadInteger Method")
End Function

'##ModelId=47B0A21800C8
Public Function ReadString(ByVal Section As String, ByVal Key As String, ByVal Default As String) As String
    On Error GoTo ReadStringErr

    '256バイトのバッファをスペースで埋める
    Dim Buffer As String * 256
    Buffer = Space$(256)
    Call GetPrivateProfileString(Section, Key, Default, Buffer, 256, IniFilePath)
    ReadString = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)

    Exit Function
ReadStringErr:
    Call RaiseError(MyUnhandledError, "IniFile:ReadString Method")
End Function

'##ModelId=47B0A21800CF
Public Function WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo WriteStringErr

    'INIファイル書き込み
    If WritePrivateProfileString(Section, Key, Value, IniFilePath) Then
        WriteString = True
    Else
        WriteString = False
    End If

    Exit Function
WriteStringErr:
    Call RaiseError(MyUnhandledError, "IniFile:WriteString Method")
End Function

#If DebugMode Then
    '##ModelId=47B0A21800D3
    Public Property Get ClassDebugID() As Long
        'ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞであれば、ｺﾝｼｭｰﾏが問合せできるこのﾌﾟﾛﾊﾟﾃｨを表に出します
        ClassDebugID = mlClassDebugID
    End Property
#End If
