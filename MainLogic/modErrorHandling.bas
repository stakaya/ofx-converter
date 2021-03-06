Attribute VB_Name = "modErrorHandling"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"47B0A21702B9"
'
Option Base 0

' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
'##ModelId=47B0A21702E9
Public Const MyObjectError1 = 1000
'##ModelId=47B0A21702F3
Public Const MyObjectError2 = 1010
'##ModelId=47B0A21702F4
Public Const MyObjectErrorN = 1234
'##ModelId=47B0A21702F5
Public Const MyUnhandledError = 9999

' This function will retrieve an error description from a resource
' file (.RES).  The ErrorNum is the index of the string
' in the resource file.  Called by RaiseError
'##ModelId=46E590E601D5
Private Function GetErrorTextFromResource(ErrorNum As Long) As String
      On Error GoTo GetErrorTextFromResourceError
      Dim strMsg As String
     
      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)

      Exit Function
GetErrorTextFromResourceError:
      If Err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
      End If
End Function

'There are a number of methods for retrieving the error
'message.  The following method uses a resource file to
'retrieve strings indexed by the error number you are
'raising.
'##ModelId=47B0A21702FE
Public Sub RaiseError(ErrorNumber As Long, Source As String)
      Dim strErrorText As String

      strErrorText = GetErrorTextFromResource(ErrorNumber)

      'raise an error back to the client
      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText
End Sub


