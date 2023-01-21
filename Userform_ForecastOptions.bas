' UserForm that allows the user to set options for the forecast '
' ------------------------------------------------------------- '

Option Explicit


' Temporarily store these ranges after the user enters them so we can do further validation when the form is submitted
Dim trainingDataRange As Range
Dim holdoutDataRange As Range


' Show a dialog box allowing the user to select a range
Private Sub showRangeSelectionDialog(range As Range, textBox As Control)
	Debug.Print("Opening range selection dialog")

	' Hide the form while the range selection dialog is open
	Me.Hide

	' Loop until the user selects a valid range
	Do While True
		On Error Resume Next

		Set range = Application.InputBox("Please select a range", Type:=8)

		On Error GoTo 0

		' Close dialog if user pressed cancel
		If range Is Nothing Then
			Debug.Print("User cancelled range selection")

			Exit Do
		ElseIf Validation.isRangeValid(range) Then
			textBox = range.Address(False, False)

			Debug.Print("Input range is valid")

			Exit Do
		End If

		' Clear the range selection because it was invalid
		Set range = Nothing
	Loop

	Me.Show
End Sub


' Enable/disable the manual smoothing section based on whether the user has selected manual smoothing
Private Sub toggleAutomaticSmoothing(isEnabled As Boolean)
	Debug.Print("Toggling automatic smoothing")

	Option_AutomaticSmoothing.Value = isEnabled
	Option_ManualSmoothing.Value = Not isEnabled
	Label_LS.Enabled = Not isEnabled
	Label_TS.Enabled = Not isEnabled
	Label_SS.Enabled = Not isEnabled
	Textbox_LS.Enabled = Not isEnabled
	Textbox_TS.Enabled = Not isEnabled
	Textbox_SS.Enabled = Not isEnabled
	Spinner_LS.Enabled = Not isEnabled
	Spinner_TS.Enabled = Not isEnabled
	Spinner_SS.Enabled = Not isEnabled
End Sub


' Increment or decrement a textbox number by the specified amount and validate it
Private Sub updateTextboxNum(textBox As Control, min As Double, max As Double, amount As Double)
	Debug.Print("Updating textbox num to " & amount)

	textBox = textBox + amount

	Call Validation.validateTextboxNum(textBox, min, max)
End Sub


' Close the Forecast Options UserForm
Private Sub closeUserForm()
	Debug.Print("Closing Forecast Options dialog")

	Unload Me
End Sub


' Reset forecast options to their default values
Private Sub resetForecastOptions(Optional showDialog As Boolean = True)
	Debug.Print("Resetting forecast options")

	Call toggleAutomaticSmoothing(False)

	Textbox_p = Constants.DEFAULT_P
	Textbox_LS = Constants.DEFAULT_LS
	Textbox_TS = Constants.DEFAULT_TS
	Textbox_SS = Constants.DEFAULT_SS
	Textbox_k = Constants.DEFAULT_K
	Checkbox_MSE = Constants.DEFAULT_INCLUDE_MSE
	Checkbox_BIAS = Constants.DEFAULT_INCLUDE_BIAS
	Checkbox_MAD = Constants.DEFAULT_INCLUDE_MAD
	Checkbox_MAPE = Constants.DEFAULT_INCLUDE_MAPE
	Checkbox_MAE = Constants.DEFAULT_INCLUDE_MAE
	Checkbox_Charts = Constants.DEFAULT_INCLUDE_CHARTS

	If showDialog Then Call showInfoMsg("Forecast options have been reset to their default values.")
End Sub


' Open the Help UserForm
Private Sub openHelpDialog()
	Debug.Print("Opening help dialog")

	Userform_Help.Show
End Sub


' Perform validation on forecast options and pass them to the Forecast module
Private Sub generateForecast()
	Debug.Print("Generating forecast")

	' Do some last minute validation before generating the forecast
	If Validation.areRangesValid(trainingDataRange, holdoutDataRange, Textbox_p, Textbox_k) Then
		Debug.Print("Forecast options are valid")

		' Create a new ForecastOption object
		Dim forecastOptions As New ForecastOptions

		Set forecastOptions.trainingDataRange = trainingDataRange
		Set forecastOptions.holdoutDataRange = holdoutDataRange

		forecastOptions.p = Textbox_p
		forecastOptions.isAutomaticSmoothing = Option_AutomaticSmoothing
		forecastOptions.LS = Textbox_LS
		forecastOptions.TS = Textbox_TS
		forecastOptions.SS = Textbox_SS
		forecastOptions.k = Textbox_k
		forecastOptions.includeMSE = Checkbox_MSE
		forecastOptions.includeBIAS = Checkbox_BIAS
		forecastOptions.includeMAD = Checkbox_MAD
		forecastOptions.includeMAPE = Checkbox_MAPE
		forecastOptions.includeMAE = Checkbox_MAE
		forecastOptions.includeCharts = Checkbox_Charts

		' Pass the forecast options to the Forecast module
		Call ForecastEngine.startForecastGeneration(forecastOptions)
		Call closeUserForm
		Call showInfoMsg("Forecast generated successfully. See the output worksheet for more details.")
	End If
End Sub


' Event handler. Called when the `Show` method is called on the UserForm
Private Sub UserForm_Initialize()
	Debug.Print ("Initializing Forecast Options UserForm")

	' Replace label text with the values we defined in the Constants module
	Userform_ForecastOptions.Caption = Constants.ADDIN_NAME & " - Forecast Options"

	Label_AddinName = Constants.ADDIN_NAME
	Label_AddinAttribution = Constants.ADDIN_ATTRIBUTION

	Frame_InputData.Caption = Constants.LBL_INPUT_DATA
	Frame_SmoothingParameters.Caption = Constants.LBL_SMOOTHING_PARAMETERS
	Frame_AutomaticSmoothing.Caption = Constants.LBL_AUTOMATIC_SMOOTHING
	Frame_ManualSmoothing.Caption = Constants.LBL_MANUAL_SMOOTHING
	Frame_OutputOptions.Caption = Constants.LBL_OUTPUT_OPTIONS
	Frame_IncludedMetrics.Caption = Constants.LBL_INCLUDED_METRICS

	Label_TrainingDataRange = Constants.LBL_TRAINING_DATA_RANGE & ":"
	Label_HoldoutDataRange = Constants.LBL_HOLDOUT_DATA_RANGE & ":"
	Label_p = Constants.LBL_P & ":"
	Label_LS = Constants.LBL_LS & ":"
	Label_TS = Constants.LBL_TS & ":"
	Label_SS = Constants.LBL_SS & ":"
	Label_k = Constants.LBL_K & ":"

	Checkbox_RMSE.Caption = Constants.LBL_RMSE
	Checkbox_MSE.Caption = Constants.LBL_MSE
	Checkbox_BIAS.Caption = Constants.LBL_BIAS
	Checkbox_MAD.Caption = Constants.LBL_MAD
	Checkbox_MAPE.Caption = Constants.LBL_MAPE
	Checkbox_MAE.Caption = Constants.LBL_MAE
	Checkbox_Charts.Caption = Constants.LBL_CHARTS

	Label_Cancel = Constants.LBL_CANCEL
	Label_Reset = Constants.LBL_RESET
	Label_Help = Constants.LBL_HELP
	Label_Forecast = Constants.LBL_FORECAST

	' Set button colors
	With Image_Cancel
		.BackColor = Constants.COLOR_RED_LIGHT
		.BorderColor = Constants.COLOR_RED_DARK
	End With

	With Image_Reset
		.BackColor = Constants.COLOR_AMBER_LIGHT
		.BorderColor = Constants.COLOR_AMBER_DARK
	End With

	With Image_Help
		.BackColor = Constants.COLOR_BLUE_LIGHT
		.BorderColor = Constants.COLOR_BLUE_DARK
	End With

	With Image_Forecast
		.BackColor = Constants.COLOR_GREEN_LIGHT
		.BorderColor = Constants.COLOR_GREEN_DARK
	End With

	' Reset the forecast options to their default values
	Call resetForecastOptions(False)
End Sub


' Event handler. Called when the textbox gets focus
Private Sub Textbox_TrainingDataRange_Enter()
	Call showRangeSelectionDialog(trainingDataRange, Textbox_TrainingDataRange)
End Sub


' Event handler. Called when the button is clicked
Private Sub Button_TrainingDataRange_Click()
	Call showRangeSelectionDialog(trainingDataRange, Textbox_TrainingDataRange)
End Sub


Private Sub Textbox_HoldoutDataRange_Enter()
	Call showRangeSelectionDialog(holdoutDataRange, Textbox_HoldoutDataRange)
End Sub


Private Sub Button_HoldoutDataRange_Click()
	Call showRangeSelectionDialog(holdoutDataRange, Textbox_HoldoutDataRange)
End Sub


Private Sub Option_AutomaticSmoothing_Click()
	Call toggleAutomaticSmoothing(True)
End Sub


Private Sub Option_ManualSmoothing_Click()
	Call toggleAutomaticSmoothing(False)
End Sub


' Event handler. Called when the user presses a key in the textbox. We can use this to to filter input in real-time
Private Sub Textbox_p_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	KeyAscii = Validation.blockNonNumericChars(KeyAscii)
End Sub


Private Sub Textbox_LS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	KeyAscii = Validation.blockNonNumericChars(KeyAscii, True)
End Sub


Private Sub Textbox_TS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	KeyAscii = Validation.blockNonNumericChars(KeyAscii, True)
End Sub


Private Sub Textbox_SS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	KeyAscii = Validation.blockNonNumericChars(KeyAscii, True)
End Sub


Private Sub Textbox_k_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
	KeyAscii = Validation.blockNonNumericChars(KeyAscii)
End Sub


' Event handler. Called when the textbox loses focus after an update
Private Sub Textbox_p_AfterUpdate()
	Call Validation.validateTextboxNum(Textbox_p, Constants.PERIODS_MIN, Constants.PERIODS_MAX)
End Sub


Private Sub Textbox_LS_AfterUpdate()
	Call Validation.validateTextboxNum(Textbox_LS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX)
End Sub


Private Sub Textbox_TS_AfterUpdate()
	Call Validation.validateTextboxNum(Textbox_TS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX)
End Sub


Private Sub Textbox_SS_AfterUpdate()
	Call Validation.validateTextboxNum(Textbox_SS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX)
End Sub


Private Sub Textbox_k_AfterUpdate()
	Call Validation.validateTextboxNum(Textbox_k, Constants.PERIODS_MIN, Constants.PERIODS_MAX)
End Sub


' Event handler. Called when the `up` button is clicked
Private Sub Spinner_p_SpinUp()
	Call updateTextboxNum(Textbox_p, Constants.PERIODS_MIN, Constants.PERIODS_MAX, Constants.PERIODS_STEP)
End Sub


' Event handler. Called when the `down` button is clicked
Private Sub Spinner_p_SpinDown()
	Call updateTextboxNum(Textbox_p, Constants.PERIODS_MIN, Constants.PERIODS_MAX, -Constants.PERIODS_STEP)
End Sub


Private Sub Spinner_LS_SpinUp()
	Call updateTextboxNum(Textbox_LS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_LS_SpinDown()
	Call updateTextboxNum(Textbox_LS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, -Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_TS_SpinUp()
	Call updateTextboxNum(Textbox_TS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_TS_SpinDown()
	Call updateTextboxNum(Textbox_TS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, -Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_SS_SpinUp()
	Call updateTextboxNum(Textbox_SS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_SS_SpinDown()
	Call updateTextboxNum(Textbox_SS, Constants.SMOOTHING_MIN, Constants.SMOOTHING_MAX, -Constants.SMOOTHING_STEP)
End Sub


Private Sub Spinner_k_SpinUp()
	Call updateTextboxNum(Textbox_k, Constants.PERIODS_MIN, Constants.PERIODS_MAX, Constants.PERIODS_STEP)
End Sub


Private Sub Spinner_k_SpinDown()
	Call updateTextboxNum(Textbox_k, Constants.PERIODS_MIN, Constants.PERIODS_MAX, -Constants.PERIODS_STEP)
End Sub


' Event handler. Called when the user clicks the image button
Private Sub Image_Cancel_Click()
	Call closeUserForm
End Sub


Private Sub Label_Cancel_Click()
	Call closeUserForm
End Sub


Private Sub Image_Reset_Click()
	Call resetForecastOptions
End Sub


Private Sub Label_Reset_Click()
	Call resetForecastOptions
End Sub


Private Sub Image_Help_Click()
	Call openHelpDialog
End Sub


Private Sub Label_Help_Click()
	Call openHelpDialog
End Sub


Private Sub Image_Forecast_Click()
	Call generateForecast
End Sub


Private Sub Label_Forecast_Click()
	Call generateForecast
End Sub
