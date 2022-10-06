Attribute VB_Name = "VBA_TIMING_FUNCTIONS"
Option Explicit

Private Declare PtrSafe Sub Kernel_Sleep_Milliseconds Lib "Kernel32.dll" Alias "Sleep" (ByVal Sleep_Milliseconds As Long)
Private Declare PtrSafe Sub Get_Local_Time Lib "Kernel32.dll" Alias "GetLocalTime" (ByRef Local_Time As Kernel_System_Time)
Private Declare PtrSafe Function QPC Lib "Kernel32.dll" Alias "QueryPerformanceCounter" (ByRef Query_PerfCounter As Currency) As Long
Private Declare PtrSafe Function QPF Lib "Kernel32.dll" Alias "QueryPerformanceFrequency" (ByRef Query_PerfFrequency As Long) As Long

Private Type Kernel_System_Time

             Year           As Integer
             Month          As Integer
             WeekDay        As Integer
             Day            As Integer
             Hour           As Integer
             Minute         As Integer
             Second         As Integer
             Milliseconds   As Integer
End Type

Private Perf_Counter As Currency
Private Const LONG_1 As Long = 1
Private Const LONG_12 As Long = 12
Private Const QPC_Adjust As Long = 1000
Private Const LONG_10000 As Long = 10000
Private Const TEXT_DOT As String = "."

Public Sub KERNEL_SLEEP(Optional Wait_Milliseconds As Long)    ' Sleeps for Wait_Milliseconds while keeping VBA responsive

Dim Wait_Expired As Boolean
Dim Wait_Remaining As Long
Dim Loop_Sleep_Time As Long
Dim Effective_Wait_Time As Long

Const Loop_Time As Long = 100  ' MilliSeconds

Wait_Remaining = IIf(Wait_Milliseconds < LONG_1, LONG_1, Wait_Milliseconds)

Effective_Wait_Time = IIf(Wait_Remaining < Loop_Time, Wait_Remaining, Loop_Time)

Do

 Wait_Expired = Wait_Remaining < LONG_1
        
    If Not Wait_Expired Then

        Loop_Sleep_Time = IIf(Wait_Remaining < Effective_Wait_Time, Wait_Remaining, Effective_Wait_Time)
            
        Kernel_Sleep_Milliseconds Loop_Sleep_Time
            
        Wait_Remaining = Wait_Remaining - Loop_Sleep_Time
            
    End If
   
 DoEvents ' prevents "VBA Not Responding"
   
Loop Until Wait_Expired

End Sub

Public Function TimeStamp() As String               ' returns VBA time with Millseconds suffix - HH.MM.SS.mmm

' Application.Volatile  ' for Excel Worksheet cell use only

Dim Timestamp_Time As Kernel_System_Time
Dim Timestamp_String As String * LONG_12            ' extends timestamp string to consistent length of 12 characters for milliseconds < 100

Get_Local_Time Timestamp_Time

Timestamp_String = Time() & TEXT_DOT & Timestamp_Time.Milliseconds

TimeStamp = Timestamp_String

End Function

Public Function GET_QPF_RESOLUTION() As Long        ' Returns QPF for this computer, typically 10,000,000 for modern PCs

' Application.Volatile  ' for Excel Worksheet cell use only

Dim QPF_Resolution As Long

QPF QPF_Resolution

GET_QPF_RESOLUTION = QPF_Resolution

End Function

Public Function QPF_TEST() As Long                  ' Checks that QPC_Adjust value in Declarations section above is correct

Dim Temp_Adjust As Long

Temp_Adjust = GET_QPF_RESOLUTION \ LONG_10000

Debug.Print "Correct QPC_Adjust value for this computer = " & Temp_Adjust
Debug.Print "Typical QPC_Adjust value for new computers = 1000"

QPF_TEST = Temp_Adjust

End Function

Public Function GET_QPC_SECONDS() As Long           ' Seconds since last system boot

' Application.Volatile  ' for Excel Worksheet cell use only

QPC Perf_Counter

GET_QPC_SECONDS = Perf_Counter \ QPC_Adjust         ' Integer division

End Function

Public Function GET_QPC_MILLISECONDS() As Long      ' MilliSeconds since last system boot

' Application.Volatile  ' for Excel Worksheet cell use only

QPC Perf_Counter

GET_QPC_MILLISECONDS = Perf_Counter \ LONG_1        ' Integer division

End Function

Public Function GET_QPC_MICROSECONDS() As Currency  ' MicroSeconds since last system boot

' Application.Volatile  ' for Excel Worksheet cell use only

QPC Perf_Counter

GET_QPC_MICROSECONDS = Int(Perf_Counter * QPC_Adjust)

End Function
