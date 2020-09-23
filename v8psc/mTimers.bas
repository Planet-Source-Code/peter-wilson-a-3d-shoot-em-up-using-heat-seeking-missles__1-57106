Attribute VB_Name = "mTimers"
Option Explicit

' The Sleep function suspends the execution of the current thread for a specified interval.
' (This is like a Pause function for the EXE, ie. It will slow the whole EXE down.)
' (1000 Milliseconds = 1 Second)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' The GetTickCount function retrieves the number of milliseconds
' that have elapsed since the system was started. This is a great API
' for measuring time intervals (much better than Timer controls!)
Public Declare Function GetTickCount Lib "kernel32" () As Long


' The following declarations are similar to GetTickCount except
' QueryPerformanceCounter returns a much more accurate result. (The reason I am using
' this kind of accuracy is so that I can get an accurate "Frames drawn per second". The good-old
' GetTickCount API wasn't accurate enough)
' The ULARGE_INTEGER structure is used to specify a 64-bit unsigned integer value.
Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As ULARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As ULARGE_INTEGER) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function convert_ULargeInt2Currency(li As ULARGE_INTEGER) As Currency
    
    ' This function converts the 64 bits of data in the ULARGE_INTEGER structure into
    ' a Visual Basic equivalent data type - the 64-bit Currency data type.
    '
    ' Internally, the VB Currency data type is actually a 64-bit integer.
    ' However, Visual Basic scales down by a factor of 10,000 to produce
    ' four digits after the decimal point. So, although the data type is
    ' a 64-bit integer, Visual Basic will display the value of any
    ' Currency-type variable as having a decimal point followed by four digits.
    ' Therefore, in order to display a 64-bit value copied into the variable correctly,
    ' you must first multiply the variable by 10,000. This will shift the decimal
    ' point four places to the right, resulting in the display of the actual value.
    
    Dim curTempValue As Currency
    
    CopyMemory curTempValue, li, 8
    
    ' Note: You don't really have to multiply by 10000 (it's optional).
    ' If you have a GHz PC then the large values may not fit anyway.
    convert_ULargeInt2Currency = curTempValue '* 10000
    
End Function
Public Function GetPerformanceFrequency() As Currency
    
    ' The QueryPerformanceFrequency function retrieves the frequency
    ' of the high-resolution performance counter, if one exists.
    ' The frequency cannot change while the system is running.
    
    Dim lngReturnValue As Long              ' Generic API return value.
    Dim temp64BitBuffer As ULARGE_INTEGER
    
    lngReturnValue = QueryPerformanceFrequency(temp64BitBuffer)
    
    If lngReturnValue = 0 Then
        ' An error has occured, or the installed hardware does not support
        ' a high-resolution performance counter.
        GetPerformanceFrequency = 0
    Else
        ' Convert from an API specific data type, to a Visual Basic one.
        ' This returns the number of Hertz the counter is running at.
        ' Below is the results from two of my computer.
        ' DELL XPS T500 PentiumIII 500MHz       :       3,579,545 (3.58MHz)
        ' DELL PowerEdge 400SC Pentium4 2.4GHz  :   2,393,840,000 (2.39GHz)
        GetPerformanceFrequency = convert_ULargeInt2Currency(temp64BitBuffer)
        
    End If
    
End Function
Public Function GetPerformanceCounter() As Currency
    
    ' The QueryPerformanceCounter function retrieves the current value
    ' of the high-resolution performance counter, if one exists.
    
    Dim lngReturnValue As Long              ' Generic API return value.
    Dim temp64BitBuffer As ULARGE_INTEGER
    
    lngReturnValue = QueryPerformanceCounter(temp64BitBuffer)
    
    If lngReturnValue = 0 Then
        ' An error has occured, or the installed hardware does not support
        ' a high-resolution performance counter.
        GetPerformanceCounter = 0
    Else
        ' Convert from an API specific data type, to a Visual Basic one.
        GetPerformanceCounter = convert_ULargeInt2Currency(temp64BitBuffer)
    End If
    
End Function

