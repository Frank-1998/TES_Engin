' UserForm that shows instructions for users '
' ------------------------------------------ '

Option Explicit


' Event handler. Called when the `Show` method is called on the UserForm
Private Sub UserForm_Initialize()
  Debug.Print ("Initializing Help UserForm")

	Dim BR As String

	BR = vbNewline & vbNewline

  Userform_Help.Caption = Constants.ADDIN_NAME & " - Help"

	Frame_InputData.Caption = Constants.LBL_INPUT_DATA
	Frame_SmoothingParameters.Caption = Constants.LBL_SMOOTHING_PARAMETERS
	Frame_OutputOptions.Caption = Constants.LBL_OUTPUT_OPTIONS

	Textbox_Header.Text = "Welcome to the " & Constants.ADDIN_NAME & ". All options are validated automatically when you enter them or click the " & Constants.LBL_FORECAST & " button."

	Textbox_InputData.Text = "Click the `Choose` buttons to select ranges for training data and holdout data. The selected ranges should be one column wide and contain only numbers (ie. don't include column headers or date columns in the selection)." & BR & _
	_
	"Make sure the " & Constants.LBL_TRAINING_DATA_RANGE & " is large enough that there is a data point available for each seasonality period (ex. if " & Constants.LBL_P & " is set to 12, make sure the " & Constants.LBL_TRAINING_DATA_RANGE & " has at least 12 rows)." & BR & _
	_
	"The " & Constants.LBL_HOLDOUT_DATA_RANGE & " is used to calculate out-of-sample error, so make sure the range does not contain blank cells and that the value for " & Constants.LBL_K & " is larger than the " & Constants.LBL_HOLDOUT_DATA_RANGE & "." & BR & _
	_
	"Enter an integer between " & Constants.PERIODS_MIN & " and " & Constants.PERIODS_MAX & " for " & Constants.LBL_P & ". This value specifies the number of seasonal periods there are in one cycle. For example, if we have 4 data points per year (quarterly), then " & Constants.LBL_P & " should be set to 4."

	Textbox_SmoothingParameters.Text = "Choose the " & Constants.LBL_AUTOMATIC_SMOOTHING & " option to have the add-in determine the optimal values for " & Constants.LBL_LS & ", " & Constants.LBL_TS & ", and " & Constants.LBL_SS & " automatically." & BR & _
	_
	"Choose the " & Constants.LBL_MANUAL_SMOOTHING & " option if you want to specify the smoothing parameters manually. Each value can range from " & Constants.SMOOTHING_MIN & " to " & Constants.SMOOTHING_MAX & "."

	Textbox_OutputOptions.Text = "Enter an integer between " & Constants.PERIODS_MIN & " and " & Constants.PERIODS_MAX & " for " & Constants.LBL_K & ". This value specifies the number of periods in the future we will generate forecast values for." & BR & _
	_
	"Select each metric you want to include in the output report. " & Constants.LBL_RMSE & " will always be included."

	Textbox_Footer.Text = "Click " & Constants.LBL_CANCEL & " to close the Forecast Options dialog, " & Constants.LBL_RESET & " to set the forecast options to their initial values, " & Constants.LBL_HELP & " to display this help dialog, and " & Constants.LBL_FORECAST & " to start generating a forecast with the given options."
End Sub
