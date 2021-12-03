Public Dim Application as ExcelApplication
Public Dim Workbooks as Workbooks
Public Dim ActiveWorkbook as Workbook
Public Dim ActiveSheet as Worksheet
Public Dim ActiveCell as Range
Public Dim Rows as Range
Public Dim Columns as Range
Public Dim Cells as Range
Public Dim Range as Range
Public Dim Selection as Range

Public Sub MsgBox(message as Object)
End Sub

Public Function Now() as ExcelDate
End Function

Public Function ToDate(ts as TimeSpan) as Date
End Function

Public Function Right(str as String, num as Integer) as String
End Function

Public Function Left(str as String, num as Integer) as String
End Function

Public Dim TestValue as Integer

Public Sub MainSetValue()
	TestValue = 123
End Sub

Public Function MainGetValue() as Integer
	MainGetValue = TestValue
End Function