VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Memento As MyMemento
 
Friend Sub SetCopy(NewCopy As MyMemento)
    Memento = NewCopy
End Sub

Public Function Copy() As clsParams
    Dim Result As clsParams
    Set Result = New clsParams
    Result.SetCopy Memento
    Set Copy = Result
End Function
