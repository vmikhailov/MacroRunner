Public Dim TestValue2 as Integer

Public Sub Main()
    Dim r As Integer
    Dim c As Integer
    Dim s As Integer
	    
	ActiveSheet.Value = 1

    Let r = 3
    For c = 2 To 60
        s = s + Cells(r, c).Value * 2
    Next
    
    Cells(r, c + 1).Value = s
	TestValue = 123
	Context.TestValue = 456

	Dim a as new ClassA
	a.FieldA = 1
End Sub

Public Class ClassA
	Public Dim FieldA as Integer
End Class
