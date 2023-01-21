# TES Forecaster Addin

by Frank, Ibrahim, John, and Yiteng


## Installation

We have made this project into an Excel add-in so that it can be used easily on any worksheet. Installation steps are as follows:


1. Unblock the `TES Forecaster Addin.xlam` file in Windows
2. Open Excel
3. Click Developer > Excel Add-ins > Browse
4. Find `TES Forecaster Addin.xlam` and make sure it is selected in the list of add-ins
5. Click `OK`
6. After installation, a button will be added under the `Add-ins` tab in Excel that can be used to launch the add-in

Uncheck `TES Forecaster Addin` in the list of add-ins to remove it again.


## Usage
Launch the `Forecast Options` dialog by clicking the button in the ribbon under Add-ins > TES Forecaster Addin. If this button does not show up for whatever reason, the add-in can be launched by running the following macro from a cell or button: `=openForecastOptionsDialog()`.

Click the `Help` button in the `Forecast Options` dialog for more details on required input.


**TODO (REMOVE THIS LATER):**
- Tab doesn't work on button because Excel keep resetting TabStop property
- Make sub to remove UI backup on uninstall
- Add error handling to ribbon code
