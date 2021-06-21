VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFileTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46B52E410135"
Option Explicit

Implements IFixture

'##ModelId=46B52E410153
Private Assert As IAssert

'##ModelId=46B52E410154
Private Sub IFixture_Setup(TestAssert As vbUnit3.IAssert)
    Set Assert = TestAssert
End Sub

'##ModelId=46B52E41015D
Private Sub IFixture_TearDown()
End Sub

'##ModelId=46B52E41015E
Public Sub TestReadInteger()
    Dim Ini As New IniFile
    Assert.LongsEqual Ini.ReadInteger("License", "Playcount", -1), 2
End Sub

'##ModelId=46B52E41015F
Public Sub TestReadString()
    Dim Ini As New IniFile
    Assert.StringsEqual Ini.ReadString("License", "Playcount", ""), "2"
End Sub

'##ModelId=46B52E410160
Public Sub TestWriteString()
    Dim Ini As New IniFile
    Assert.VariantsEqual Ini.WriteString("License", "Playcount", "2"), True
End Sub

