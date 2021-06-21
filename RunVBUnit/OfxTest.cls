VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OfxTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46B52E410045"
Option Explicit

Implements IFixture

'##ModelId=46B52E4100BD
Private Assert As IAssert

Private Sub IFixture_Setup(TestAssert As vbUnit3.IAssert)
    Set Assert = TestAssert
End Sub

Private Sub IFixture_TearDown()
End Sub

Public Sub TestSetOutputFile()
End Sub

