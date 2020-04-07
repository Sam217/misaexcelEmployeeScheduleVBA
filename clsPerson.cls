VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
Private m_id As Integer
Private m_count(5) As Integer
Public holidayWorks As Integer

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(aName As String)
    m_name = aName
End Property

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let Id(aId As Integer)
    m_id = aId
End Property

Public Property Get count(idx As Integer) As Integer
    count = m_count(idx)
End Property

Public Property Let count(idx As Integer, aCount As Integer)
    m_count(idx) = aCount
End Property

Public Sub Init()
    m_id = 0
    m_count(5) = 0
    ReDim m_counts(4)
    holidayWorks = 0
    m_count(0) = 0
    m_count(1) = 0
    m_count(2) = 0
    m_count(3) = 0
    m_count(4) = 0
End Sub

Private Sub Class_Initialize()
    Init
End Sub

