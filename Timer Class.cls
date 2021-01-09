VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'
' API Declarations
'
Private Declare Function QueryPerformanceCounter Lib "kernel32" _
(lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" _
(lpFrequency As Currency) As Long

Private Declare Function GetTickCount Lib "kernel32" () _
As Long


'
' Local Variables
'
Dim timStart As Currency    ' The Starting Counter of the Timer
Dim timElapsed As Double    ' No. of Elapsed Milliseconds
Dim nFrequency As Currency  ' Frequency of the Hi-Resolution timer if it is avilable, else zero
Dim bRunning As Boolean     ' Is the Timer running ?

Private Sub Class_Initialize()
'
' Check if Hi-resolution timer is available
'
Dim APIRetVal As Long

APIRetVal = QueryPerformanceFrequency(nFrequency)

If APIRetVal = 0 Then
    nFrequency = 0          ' Zero indicates that the ordinary GetTickCount function must be used
End If
End Sub

Public Property Get IsHiTimer() As Boolean
Attribute IsHiTimer.VB_Description = "Returns True if the System supports a Hi-Resolution timer"
'
' Returns True if a Hi-Performance timer is used
'
IsHiTimer = (nFrequency <> 0)
End Property

Public Sub StartTimer()
Attribute StartTimer.VB_Description = "Starts the Timer"
'
' Starts the Counter
'
timElapsed = 0      ' Reset Counter
bRunning = True

If nFrequency <> 0 Then
    QueryPerformanceCounter timStart
Else
    timStart = GetTickCount
End If
End Sub

Public Sub EndTimer()
Attribute EndTimer.VB_Description = "Ends the Timer"
'
' Ends the Counter
'
timElapsed = GetElapsedTime     ' Store this Value
bRunning = False
End Sub

Public Property Get Elapsed() As Double
Attribute Elapsed.VB_Description = "Returns Elapsed Time in MilliSeconds. If the Timer is running, returns Zero."
'
' Returns Elapsed time
'
If bRunning Then
    Elapsed = GetElapsedTime    ' Return Intermediate Time
Else
    Elapsed = timElapsed        ' Return the Last counter's timing
End If
End Property

Private Function GetElapsedTime() As Double
'
' Calculates the Elapsed Time
'
Dim timEnd As Currency

If nFrequency <> 0 Then
    QueryPerformanceCounter timEnd
    
    ' Get Elasped Time : Result is in Seconds, so multiply by 1000 to get result in msec
    GetElapsedTime = ((timEnd - timStart) / nFrequency) * 1000
Else
    timEnd = GetTickCount
    GetElapsedTime = timEnd - timStart
End If
End Function
