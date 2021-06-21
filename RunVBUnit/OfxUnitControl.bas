Attribute VB_Name = "OfxUnitControl"
Option Explicit

Public Sub Main()
    Dim Runner As New TestRunner
    Dim Client As ISuite

    Set Client = CreateObject("OfxTestRunner.OfxAllTests")

    Runner.AddSuite Client
    Runner.Run True, True
End Sub

