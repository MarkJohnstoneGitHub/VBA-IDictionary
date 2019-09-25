Attribute VB_Name = "modMicroTimer"
'@Folder("VBA-IScriptingDictionary.Tests.Performance")

'See:
'https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff700515(v=office.14)#Office2007excelPerf_MakingWorkbooksCalculateFaster
'https://codereview.stackexchange.com/questions/67596/a-lightning-fast-stringbuilder

Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" _
        Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" _
        Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


Function MicroTimer() As Double
'

' Returns seconds.
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0

' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency

' Get ticks.
    getTickCount cyTicks1

' Seconds
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function
