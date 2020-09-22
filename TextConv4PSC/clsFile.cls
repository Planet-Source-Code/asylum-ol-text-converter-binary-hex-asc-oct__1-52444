VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Import(ByVal File As String) As String

Dim Buffer As String
    Close #1
    Open File For Binary Access Read As #1 '// you could use input in a loop, but that's boring =P
    
    Buffer = Space$(LOF(1))
    Get #1, , Buffer
    
    Import = Buffer
End Function
