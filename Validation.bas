' Module for validating user input '
' -------------------------------- '

Option Explicit
Option Private Module


' Show an error message if the user enters an invalid value for range
Private Sub showInvalidRangeErrorMsg(range As Range, msg As String)
	Call showErrorMsg("Invalid range: " & range.Address(false, false) & vbNewLine & vbNewline & msg)
End Sub


' Prevent the user from entering invalid characters into a numeric textbox
' Adapted from https://stackoverflow.com/a/41770674/1378560
Public Function blockNonNumericChars(ByVal KeyAscii As MSForms.ReturnInteger, Optional doAllowPeriod As Boolean = False) As Long
	' If doAllowPeriod is true, allow the user to enter ASCII character 46 (period)
	If (KeyAscii >= 48 And KeyAscii <= 57) Or (doAllowPeriod And KeyAscii = 46) Then
		blockNonNumericChars = KeyAscii
	Else
		blockNonNumericChars = 0
	End If
End Function


' Ensure that a textbox number is within the specified range
Public Sub validateTextboxNum(textBox As Control, min As Double, max As Double)
	Debug.Print("Validating textbox num")

	' Clip the value to the specified range
	If textBox < min Then
		textBox = min
	ElseIf textBox > max Then
		textBox = max
	End If
End Sub


' Check whether a single range is valid or not after the user specifies it
Public Function isRangeValid(range As Range) As Boolean
	Debug.Print("Validating range")

	isRangeValid = False

	' Ensure the range is a single column and contains numbers only
	If range.Columns.Count <> 1 Then
		Call showInvalidRangeErrorMsg(range, "Please select a single column of data only.")
	ElseIf WorksheetFunction.Count(range) <> Range.Count Then
		Call showInvalidRangeErrorMsg(range, "Please select a range that contains numbers only.")
	Else
		isRangeValid = True
	End If
End Function


' Check whether the training and holdout data ranges valid or not when the user tries to submit them in a form
Public Function areRangesValid(trainingDataRange As Range, holdoutDataRange As Range, p As Long, k As Long) As Boolean
	Debug.Print("Validating ranges")

	areRangesValid = False

	If trainingDataRange Is Nothing Then
		Call showErrorMsg("No training data specified. Please select a range of training data.")
	ElseIf holdoutDataRange Is Nothing Then
		Call showErrorMsg("No holdout data specified. Please select a range of holdout data.")
	ElseIf trainingDataRange.Column <> holdoutDataRange.Column Then
		Call showErrorMsg("The training data and the holdout data are in different columns. Please select two ranges that are in the same column.")
	ElseIf trainingDataRange.Rows(trainingDataRange.Rows.Count).Row + 1 <> holdoutDataRange.Row Then
		Call showErrorMsg("The training data range and the holdout data range are not adjacent. Please select two adjacent ranges.")
	ElseIf trainingDataRange.Rows.Count < p Then
		Call showErrorMsg("There are " & p & " " & Constants.LBL_P & ", but the training data range only has " & trainingDataRange.Rows.Count & " rows." & vbNewLine & vbNewline & "Please increase the number of rows in the training data range, or decrease the number of " & Constants.LBL_P & ".")
	ElseIf holdoutDataRange.Rows.Count > k Then
		Call showErrorMsg("The holdout data range has " & holdoutDataRange.Rows.Count & " rows, but we are only generating forecasts for " & k & " periods in the future." & vbNewLine & vbNewline &  "Please increase the number of " & Constants.LBL_K & " or decrease the number of holdout data rows.")
	Else
		areRangesValid = True
	End If
End Function
