' Module for adding/removing a custom group from the Excel ribbon '
' --------------------------------------------------------------- '

Option Explicit
Option Private Module


Dim fileSysObj As Object
Dim excelUiConfigPath As String
Dim excelUiBackupConfigPath As String


' Compute filenames
Private Sub setExcelUiConfigVars()
  Dim excelUiConfigDir As String

  Set fileSysObj = CreateObject("Scripting.FileSystemObject")

  excelUiConfigDir = "C:\Users\" & Environ("Username") & "\AppData\Local\Microsoft\Office\"
  excelUiConfigPath = excelUiConfigDir & EXCEL_UI_CONFIG_FILENAME
  excelUiBackupConfigPath = excelUiConfigDir & EXCEL_UI_BACKUP_CONFIG_FILENAME
End Sub


' Add a custom group to the Excel ribbon
' Adapted from https://stackoverflow.com/a/30893395/1378560
Public Sub InjectCustomRibbonTab()
  Debug.Print("Injecting custom ribbon tab")

  Call setExcelUiConfigVars

  ' Backup the original ribbon file
  Call fileSysObj.CopyFile(excelUiConfigPath, excelUiBackupConfigPath, True)

  Dim file As Long
  Dim xmlString As String

  file = FreeFile
  xmlString = "<customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><qat/><tabs><tab idMso='TabAddIns'><group id='tfaGroup' label='" & Constants.ADDIN_NAME & "' autoScale='true'><button id='openForecastOptionsDialogButton' label='" & Constants.LBL_RIBBON_BUTTON & "' size='large' imageMso='ForecastInsert' onAction='RibbonManagement.openForecastOptionsDialog'/></group></tab></tabs></ribbon></customUI>"
  xmlString = Replace(xmlString, """", "")

  Open excelUiConfigPath For Output Access Write As file
  Print #file, xmlString
  Close file
End Sub


' Restore the original Excel ribbon configuration
Public Sub ResetRibbonTabs()
  Debug.Print("Removing custom ribbon tab")

  Call setExcelUiConfigVars

  ' Restore the original ribbon file
  Call fileSysObj.CopyFile(excelUiBackupConfigPath, excelUiConfigPath, True)
End Sub


' Entry point for the ribbon button. Don't change the name of this sub. This is a function so that it can be called as a cell formula
Public Function openForecastOptionsDialog()
	Debug.Print("Opening Forecast Options dialog")

	Userform_ForecastOptions.Show
End Function
