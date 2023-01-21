' Class for storing user-provided forecast options '
' ------------------------------------------------ '

Option Explicit


' Properties
Private m_trainingDataRange As Range
Private m_holdoutDataRange As Range
Private m_p As Long
Private m_isAutomaticSmoothing As Boolean
Private m_LS As Double
Private m_TS As Double
Private m_SS As Double
Private m_k As Long
Private m_includeMSE As Boolean
Private m_includeBIAS As Boolean
Private m_includeMAD As Boolean
Private m_includeMAPE As Boolean
Private m_includeMAE As Boolean
Private m_includeCharts As Boolean


' Constructor
Private Sub Class_Initialize()
	Debug.Print("ForecastOptions object created")
End Sub


' Getters
Property Get trainingDataRange() As Range
	Set trainingDataRange = m_trainingDataRange
End Property

Property Get holdoutDataRange() As Range
	Set holdoutDataRange = m_holdoutDataRange
End Property

Property Get p() As Long
	p = m_p
End Property

Property Get isAutomaticSmoothing() As Boolean
	isAutomaticSmoothing = m_isAutomaticSmoothing
End Property

Property Get LS() As Double
	LS = m_LS
End Property

Property Get TS() As Double
	TS = m_TS
End Property

Property Get SS() As Double
	SS = m_SS
End Property

Property Get k() As Long
	k = m_k
End Property

Property Get includeMSE() As Boolean
	includeMSE = m_includeMSE
End Property

Property Get includeBIAS() As Boolean
	includeBIAS = m_includeBIAS
End Property

Property Get includeMAD() As Boolean
	includeMAD = m_includeMAD
End Property

Property Get includeMAPE() As Boolean
	includeMAPE = m_includeMAPE
End Property

Property Get includeMAE() As Boolean
	includeMAE = m_includeMAE
End Property

Property Get includeCharts() As Boolean
	includeCharts = m_includeCharts
End Property


' Setters
Property Set trainingDataRange(ByVal value As Range)
	Set m_trainingDataRange = value
End Property

Property Set holdoutDataRange(ByVal value As Range)
	Set m_holdoutDataRange = value
End Property

Property Let p(ByVal value As Long)
	m_p = value
End Property

Property Let isAutomaticSmoothing(ByVal value As Boolean)
	m_isAutomaticSmoothing = value
End Property

Property Let LS(ByVal value As Double)
	m_LS = value
End Property

Property Let TS(ByVal value As Double)
	m_TS = value
End Property

Property Let SS(ByVal value As Double)
	m_SS = value
End Property

Property Let k(ByVal value As Long)
	m_k = value
End Property

Property Let includeMSE(ByVal value As Boolean)
	m_includeMSE = value
End Property

Property Let includeBIAS(ByVal value As Boolean)
	m_includeBIAS = value
End Property

Property Let includeMAD(ByVal value As Boolean)
	m_includeMAD = value
End Property

Property Let includeMAPE(ByVal value As Boolean)
	m_includeMAPE = value
End Property

Property Let includeMAE(ByVal value As Boolean)
	m_includeMAE = value
End Property

Property Let includeCharts(ByVal value As Boolean)
	m_includeCharts = value
End Property
