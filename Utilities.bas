' Module for utilities used in multiple modules '
' --------------------------------------------- '

Option Explicit
Option Private Module


' Log an info message to the console and display a message box to the user
Public Sub showInfoMsg(msg As String)
	Debug.Print("[INFO]: " + msg)
	MsgBox msg, vbInformation, "Info"
End Sub


' Log an error message to the console and display a message box to the user
Public Sub showErrorMsg(msg As String)
	Debug.Print("[ERROR]: " + msg)
	MsgBox msg, vbCritical, "Error"
End Sub
