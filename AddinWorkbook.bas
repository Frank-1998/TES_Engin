' Event handlers for the default workbook '
' --------------------------------------- '

Option Explicit


Private Sub Workbook_AddInInstall()
	Call Utilities.showInfoMsg(ThisWorkbook.Name & " add-in has been installed. Run the add-in via the Add-ins tab in the ribbon.")
End Sub


Private Sub Workbook_AddinUninstall()
	Call RibbonManagement.ResetRibbonTabs
	Call Utilities.showInfoMsg(ThisWorkbook.Name & " add-in has been uninstalled. Please restart Excel to remove the button in the ribbon.")
End Sub


Private Sub Workbook_Open()
  Debug.Print("Loading addin workbook")

  Call RibbonManagement.InjectCustomRibbonTab
End Sub


Private Sub Workbook_BeforeClose(Cancel as Boolean)
  Debug.Print("Unloading addin workbook")

  Call RibbonManagement.ResetRibbonTabs
End Sub
