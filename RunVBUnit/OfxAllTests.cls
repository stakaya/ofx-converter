VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OfxAllTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46B52E4100E5"
Option Explicit

Implements ISuite

'##ModelId=46B52E4100F9
Private Function ISuite_Suite() As vbUnit3.ITest
Dim Suite As New TestSuite
    Suite.SuiteName = "全てのテスト"

    Suite.AddFixture New LoggerTest
    Suite.AddFixture New IniFileTest
    Suite.AddFixture New OfxTest

    Set ISuite_Suite = Suite
End Function


