Function LoadEmployees(ByVal target As Range) As Collection
    Dim empCol As New Collection
    Dim emp As clsPerson
    
    Dim rg As Range
    Set rg = target.CurrentRegion
    For i = 2 To rg.Columns.Count
        Set emp = New clsPerson
        emp.Name = rg.Cells(1, i)
        emp.Count = rg.Cells(2, i)
        emp.Id = i - 1
        
        empCol.Add emp
    Next
    Set LoadEmployees = empCol
End Function

Sub calcSchedule(ByVal target As Range)
    'Variables
    Dim emps(1 To 3) As Long
    Dim empsMin As Long, minIndex As Long, rndval As Long
    Dim arrIdx() As Long
    Dim var As Variant
    Dim empColl As Collection, empsEqual As New Collection
    
    Dim emplist As Range
    'here we must specify the cell where table with employees starts, in our test case "O1"
    Set emplist = Range("A1")
    Set empColl = LoadEmployees(emplist)
    ReDim arrIdx(empColl.Count - 1)
    'Code
    'Randomize
        'For i = 1 To empColl.Count
         '   rndval = Int(empColl.Count * Rnd + 1)
          '  var = Application.Match(rndval, arrIdx, 0)
           ' While Not IsError(var)
            '    rndval = Int(empColl.Count * Rnd + 1)
             '   var = Application.Match(rndval, arrIdx, 0)
            'Wend
           ' arrIdx(i) = rndval
        'Next
        
        For j = 0 To 364 'empColl.Count
            target.Offset(j, 0).ClearContents
        Next
        For i = 0 To 364 'empColl.Count
            If i < empColl.Count Then
                Dim emp As clsPerson, empFirst As clsPerson
                empsMin = empColl(1).Count
                minIndex = 1
                arrIdx(i) = minIndex
                Set empFirst = empColl(1)
                'If (i Mod 3) = 0 Then
                    'minIndex = arrIdx(1)
                'ElseIf (i Mod 3) = 1 Then
                    'minIndex = Int(2 * Rnd + 2)
                    'minIndex = arrIdx(2)
                'Else
                    'empsMin = WorksheetFunction.Min(emps)
                    'minIndex = WorksheetFunction.Match(empsMin, emps, 0)
                    For Each emp In empColl
                        If (emp.Count < empsMin) And Not (emp Is empFirst) Then
                            empsMin = emp.Count
                            minIndex = emp.Id
                            arrIdx(i) = minIndex
                        ElseIf emp.Count = empsMin And Not (emp Is empFirst) Then
                        'if any other employee has the same count, we choose randomly, who gets the next shift
                            Dim rand As Long
                            Randomize
                            rand = Round(Rnd)
                            If rand = 1 Then
                                minIndex = emp.Id
                                arrIdx(i) = minIndex
                            End If
                        End If
                    Next
            Else
                Dim k As Integer
                k = i Mod empColl.Count
                minIndex = arrIdx(k)
            End If
                
                For Each emp In empColl
                    If emp.Id = minIndex Then
                        emp.Count = emp.Count + 1
                        'target here is a cell selected by us, recently "C3"
                        target.Offset(i, 0).Value = emp.Name
                        If i Mod empColl.Count = 0 Then
                            target.Offset(i, 0).Font.Bold = True
                        End If
                        emplist.Cells(2, minIndex + 1).Value = emp.Count
                        Exit For
                    End If
                Next
        Next
        For Each emp In empColl
            Dim idx As Variant
            idx = ""
            For l = 0 To UBound(arrIdx)
                If emp.Id = arrIdx(l) Then
                    If idx = "" Then
                        idx = l + 1
                    Else
                        idx = idx & "," & l + 1
                    End If
                End If
            Next
            emplist.Cells(3, emp.Id + 1).Clear
            emplist.Cells(3, emp.Id + 1).Value = idx
        Next
End Sub


Private m_name As String
Private m_id As Integer
Private m_count As Integer

class clsPerson {
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

Public Property Get Count() As Integer
    Count = m_count
End Property

Public Property Let Count(aCount As Integer)
    m_count = aCount
End Property

Public Sub Init()
    m_id = 0
    m_count = 0
End Sub

Private Sub Class_Initialize()
    Init
End Sub
}


