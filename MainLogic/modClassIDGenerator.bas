Attribute VB_Name = "modClassIDGenerator"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"47B0A2170308"
'
Option Base 0

'##ModelId=47B0A2170326
Public Function GetNextClassDebugID() As Long
    'class ID generator
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
End Function

