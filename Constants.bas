' Module for storing constant variables used in multiple modules '
' -------------------------------------------------------------- '

Option Explicit
Option Private Module


' Addin info
Public Const ADDIN_NAME As String = "TES Forecaster Addin"
Public Const ADDIN_ATTRIBUTION As String = "by Frank, Ibrahim, John, and Yiteng"

' UI Labels
Public Const LBL_RIBBON_BUTTON As String = "Generate Forecast"
Public Const LBL_INPUT_DATA As String = "Input Data"
Public Const LBL_SMOOTHING_PARAMETERS As String = "Smoothing Parameters"
Public Const LBL_AUTOMATIC_SMOOTHING As String = "Automatic"
Public Const LBL_MANUAL_SMOOTHING As String = "Manual"
Public Const LBL_OUTPUT_OPTIONS As String = "Output Options"
Public Const LBL_INCLUDED_METRICS As String = "Included Metrics"
Public Const LBL_TRAINING_DATA_RANGE As String = "Training Data Range"
Public Const LBL_HOLDOUT_DATA_RANGE As String = "Holdout Data Range"
Public Const LBL_P As String = "Periods per Cycle (p)"
Public Const LBL_LS As String = "Level Smoothing (LS)"
Public Const LBL_TS As String = "Trend Smoothing (TS)"
Public Const LBL_SS As String = "Seasonality Smoothing (SS)"
Public Const LBL_K As String = "Forecast Periods (k)"
Public Const LBL_RMSE As String = "Root Mean Squared Error (RMSE)"
Public Const LBL_MSE As String = "Mean Squared Error (MSE)"
Public Const LBL_BIAS As String = "Bias (BIAS)"
Public Const LBL_MAD As String = "Mean Absolute Deviation (MAD)"
Public Const LBL_MAPE As String = "Means Absolute Percentage Error (MAPE)"
Public Const LBL_MAE As String = "Maximum Absolute Error (MAE)"
Public Const LBL_CHARTS As String = "Charts"
Public Const LBL_CANCEL As String = "Cancel"
Public Const LBL_RESET As String = "Reset"
Public Const LBL_HELP As String = "Help"
Public Const LBL_FORECAST As String = "Forecast"

' Colors
Public Const COLOR_RED_LIGHT As Long = &HEEEBFF&		' MD Red 50
Public Const COLOR_RED_DARK As Long = &H2828C6&			' MD Red 800
Public Const COLOR_AMBER_LIGHT As Long = &HE1F8FF&	' MD Amber 50
Public Const COLOR_AMBER_DARK As Long = &H008FFF&		' MD Amber 800
Public Const COLOR_BLUE_LIGHT As Long = &HFDF2E3&		' MD Blue 50
Public Const COLOR_BLUE_DARK As Long = &HC06515&		' MD Blue 800
Public Const COLOR_GREEN_LIGHT As Long = &HE9F8F1&	' MD Light Green 50
Public Const COLOR_GREEN_DARK As Long = &H2F8B55&		' MD Light Green 800

' Default values
Public Const DEFAULT_P As Long = 12
Public Const DEFAULT_LS As Double = 0.5
Public Const DEFAULT_TS As Double = 0.5
Public Const DEFAULT_SS As Double = 0.5
Public Const DEFAULT_K As Long = 12
Public Const DEFAULT_INCLUDE_MSE As Boolean = False
Public Const DEFAULT_INCLUDE_BIAS As Boolean = False
Public Const DEFAULT_INCLUDE_MAD As Boolean = False
Public Const DEFAULT_INCLUDE_MAPE As Boolean = False
Public Const DEFAULT_INCLUDE_MAE As Boolean = False
Public Const DEFAULT_INCLUDE_CHARTS As Boolean = False

' Input restrictions
Public Const SMOOTHING_MIN As Double = 0
Public Const SMOOTHING_MAX As Double = 1
Public Const SMOOTHING_STEP As Double = 0.1
Public Const PERIODS_MIN As Long = 1
Public Const PERIODS_MAX As Long = 999
Public Const PERIODS_STEP As Long = 1

Public Const OUTPUT_SHEET_NAME As String = "TES Forecast Report"

Public Const EXCEL_UI_CONFIG_FILENAME As String = "Excel.officeUI"
Public Const EXCEL_UI_BACKUP_CONFIG_FILENAME As String = "Excel.backup.officeUI"
