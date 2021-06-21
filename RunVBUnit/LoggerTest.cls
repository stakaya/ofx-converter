VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoggerTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46B52E410103"
Option Explicit

Implements IFixture

'##ModelId=46B52E410118
Private Sub IFixture_Setup(TestAssert As vbUnit3.IAssert)
End Sub

'##ModelId=46B52E41011A
Private Sub IFixture_TearDown()
End Sub

'##ModelId=46B52E410121
Public Sub TestLog()
    Dim log As New Logger
    log.Name = "log name"
    log.Trace ("‚ ‚ ‚ ")
    log.dbug ("‚ ‚ ‚ ")
    log.info ("‚ ‚ ‚ ")
    log.warn ("‚ ‚ ‚ ")
    log.Error ("‚ ‚ ‚ ")
    log.fatal ("‚ ‚ ‚ ")
    log.Trace ("‚ ‚ ‚ ")
    log.dbug ("‚ ‚ ‚ ")
    log.info ("‚ ‚ ‚ ")
    log.warn ("‚ ‚ ‚ ")
    log.Error ("‚ ‚ ‚ ")
    log.fatal ("‚ ‚ ‚ ")
    log.Trace ("‚ ‚ ‚ ")
    log.dbug ("‚ ‚ ‚ ")
    log.info ("‚ ‚ ‚ ")
    log.warn ("‚ ‚ ‚ ")
    log.Error ("‚ ‚ ‚ ")
    log.fatal ("‚ ‚ ‚ ")
End Sub


