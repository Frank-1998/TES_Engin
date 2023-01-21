' Module for generating forecasts '
' ------------------------------- '

Option Explicit
Option Private Module


' TODO: Implement forecast generation here
Public Sub startForecastGeneration(forecastOptionsObj As ForecastOptions)
	Debug.Print("Starting forecast generation")

	' Add a new worksheet to the end of the workbook
	Dim ws As Worksheet
	Set ws = ActiveWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))

	' TODO: Show error if this name already exists
	ws.Name = Constants.OUTPUT_SHEET_NAME

	' TODO: Example of looping over training data range
	Dim currentCell As Range

	For Each currentCell In forecastOptionsObj.trainingDataRange
		Debug.Print currentCell.Value
	Next currentCell
End Sub
