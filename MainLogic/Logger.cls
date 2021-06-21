VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"47B0A21703A8"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
Option Base 0

Option Explicit


' これを 0 に設定すると、このｸﾗｽのﾃﾞﾊﾞｯｸﾞｺｰﾄﾞが無効になります
#Const DebugMode = 1
#If DebugMode Then
    'で生成された連続するｸﾗｽIDを保持する局所変数 Class_Initialize
    '##ModelId=47B0A21703E4
    Private mlClassDebugID As Long
#End If

'##ModelId=47B0A21703E5
Private Const ALL_LEVEL = 0
'##ModelId=47B0A2180006
Private Const TRACE_LEVEL = 1
'##ModelId=47B0A2180007
Private Const DBUG_LEVEL = 2
'##ModelId=47B0A2180010
Private Const INFO_LEVEL = 3
'##ModelId=47B0A2180011
Private Const WARN_LEVEL = 4
'##ModelId=47B0A218001A
Private Const FATAL_LEVEL = 5
'##ModelId=47B0A218001B
Private Const OFF_LEVEL = 6

'##ModelId=47B0A2180024
Private LogName As String

'##ModelId=47B0A2180025
Private LogLevel As Integer


'##ModelId=47B0A2180026
Private Sub Class_Terminate()
    #If DebugMode Then
    'ｸﾗｽは破棄されます。
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=47B0A218002E
Private Sub Class_Initialize()
    #If DebugMode Then
        '次の有効なｸﾗｽIDを取得し、このｸﾗｽがうまく生成されたことを出力します。
        '
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'##ModelId=47B0A218002F
Public Sub Trace(LogMessage As String)
    On Error GoTo traceErr

    Call WriteLog("TRACE", LogMessage)

    Exit Sub
traceErr:
    Call RaiseError(MyUnhandledError, "Logger:Trace Method")
End Sub

'##ModelId=47B0A2180031
Public Sub Dbug(LogMessage As String)
    On Error GoTo dbugErr

    Call WriteLog("DEBUG", LogMessage)

    Exit Sub
dbugErr:
    Call RaiseError(MyUnhandledError, "Logger:Dbug Method")
End Sub

'##ModelId=47B0A2180038
Public Sub Info(LogMessage As String)
    On Error GoTo infoErr

    Call WriteLog("INFO", LogMessage)

    Exit Sub
infoErr:
    Call RaiseError(MyUnhandledError, "Logger:Info Method")
End Sub

'##ModelId=47B0A218003A
Public Sub Warn(LogMessage As String)
    On Error GoTo warnErr

    Call WriteLog("WARN", LogMessage)

    Exit Sub
warnErr:
    Call RaiseError(MyUnhandledError, "Logger:Warn Method")
End Sub

'##ModelId=47B0A2180042
Public Sub Error(LogMessage As String)
    On Error GoTo errorErr

    Call WriteLog("ERROR", LogMessage)

    Exit Sub
errorErr:
    Call RaiseError(MyUnhandledError, "Logger:Error Method")
End Sub

'##ModelId=47B0A2180044
Public Sub Fatal(LogMessage As String)
    On Error GoTo fatalErr

    Call WriteLog("FATAL", LogMessage)

    Exit Sub
fatalErr:
    Call RaiseError(MyUnhandledError, "Logger:Fatal Method")
End Sub

'##ModelId=47B0A2180046
Public Sub SetLogLevel(Level As Integer)
    On Error GoTo SetLogLevelErr

    LogLevel = Level

    Exit Sub
SetLogLevelErr:
    Call RaiseError(MyUnhandledError, "Logger:SetLogLevel Method")
End Sub

'##ModelId=47B0A218004E
Public Property Get name() As String
    name = LogName
End Property

'##ModelId=47B0A218004C
Public Property Let name(ByVal NewValue As String)
    LogName = NewValue
End Property

'##ModelId=47B0A218004F
Private Sub WriteLog(LogType As String, LogMessage As String)

    Dim FileNum As Integer
    Dim DateTime As String
    
    DateTime = CStr(Now())
    LogMessage = DateTime & " [" & LogType & "] - " & LogMessage
    
    Debug.Print LogMessage
    FileNum = FreeFile
    Open Format(Now(), "yyyymmdd") & ".log" For Append As #FileNum
    Print #FileNum, LogMessage
    Close #FileNum
End Sub


#If DebugMode Then
    '##ModelId=47B0A2180058
    Public Property Get ClassDebugID() As Long
        'ﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞであれば、ｺﾝｼｭｰﾏが問合せできるこのﾌﾟﾛﾊﾟﾃｨを表に出します
        ClassDebugID = mlClassDebugID
    End Property
#End If
