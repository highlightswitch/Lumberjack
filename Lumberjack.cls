VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lumberjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''
'       Declarations
'''''''''''''''''''''''''''''

' These are used to get the current time and date
Private Type systemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private st As systemTime
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As systemTime)

' A log entry is stored every time the user calls the 'log' subroutine
Private Type logEntry
    timestamp() As Integer
    description As String
End Type
Private logEntries(2047) As logEntry

Private logStartTime() As Integer 'Used to store when the 'logStart' subroutine is called
Private currLogRow As Integer 'Used to keep track of where the next log should be stored in the array 'logEntries'

' This function provides the current date and time in the form of an array of integers
Private Function dateTimeNow() As Integer()
    Dim arr(7) As Integer
    GetSystemTime st
    arr(0) = st.wYear
    arr(1) = st.wMonth
    arr(2) = st.wDayOfWeek
    arr(3) = st.wDay
    arr(4) = st.wHour
    arr(5) = st.wMinute
    arr(6) = st.wSecond
    arr(7) = st.wMilliseconds
    dateTimeNow = arr
End Function





'''''''''''''''''''''''''''''
' Public Logging Subroutines
'''''''''''''''''''''''''''''

' This subroutine should be called by the user when the want to note the start time of the log
Public Sub logStart()
    currLogRow = 0
    log ("Log Start")
End Sub

' This subroutine should be used by the user whenever they want create a lof entry with a given description 'str'
' TODO: throw error if log not started.
' TODO: handle blank string as parameter
' TODO: flagging for outputting to debug as the program progresses
' TODO: ensure the data is saved if program crashes, either to file, or to debug log
Public Sub log(str As String)

    Dim logTime() As Integer
    logTime = dateTimeNow
    
    Dim entry As logEntry
    entry.timestamp = logTime
    entry.description = str
    
    logEntries(currLogRow) = entry
    
    currLogRow = currLogRow + 1

End Sub






'''''''''''''''''''''''''''''
' Public Output Subroutines
'''''''''''''''''''''''''''''


' This subroutine should be used by the user when the want to output the current log to the debug log.
Public Sub outputToDebug()

    Debug.Print "-----------------------------------------------------"
    Debug.Print "| Timestamp               | Seconds Since Start | Seconds Since Last Log | Description"

    Dim i As Integer
    Dim leftToPrint As Boolean
    Dim str As String
    Dim entryArr() As String
    
    i = 0
    leftToPrint = logEntries(i).description <> "" 'TODO: should probably be <> null. need to test
    
    Do While leftToPrint
        
        entryArr = getOutputforEntry(i)
        str = "| " & _
            entryArr(0) & " | " & _
            entryArr(1) & "               | " & _
            entryArr(2) & "                  | " & _
            entryArr(3)
        Debug.Print str
        
        i = i + 1
        If logEntries(i).description = "" Then
            leftToPrint = False
        End If
        
    Loop
    
    Debug.Print "End"

End Sub


' This subroutine should be used by the user when the want to output the current log a specified CSV file.
' If the user specifies no CSV file, then a file will be created on the user's Desktop
Sub outputToCsvFile(Optional filePath)
    
    'If no argument provided for filePath, then set to user's Desktop
    If IsMissing(filePath) Then
        filePath = CreateObject("WScript.Shell").specialfolders("Desktop") & "\LumberjackLog.csv"
    End If

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set textStream = fs.CreateTextFile(filePath, True)
    textStream.WriteLine ("Timestamp,Seconds Since Start,Seconds Since Last Log,Description")
    
    
    Dim i As Integer
    Dim leftToPrint As Boolean
    Dim str As String
    Dim entryArr() As String
    
    i = 0
    leftToPrint = logEntries(i).description <> "" 'TODO: should probably be <> null. need to test
    
    Do While leftToPrint
        
        entryArr = getOutputforEntry(i)
        str = entryArr(0) & "," & _
            entryArr(1) & "," & _
            entryArr(2) & "," & _
            entryArr(3)
        textStream.WriteLine (str)
        
        i = i + 1
        If logEntries(i).description = "" Then
            leftToPrint = False
        End If
        
    Loop
    
    textStream.Close

End Sub






'''''''''''''''''''''''''''''
' Private Output Functions
'''''''''''''''''''''''''''''

Private Function getOutputforEntry(i As Integer) As String()
    Dim arr(3) As String
    
    arr(0) = timeArrToString(logEntries(i).timestamp)
    
    'If i=0, then this is the first entry. We must put '0.000' for elements 1 and 2
    If i = 0 Then
        arr(1) = "0.000"
        arr(2) = "0.000"
    Else
        arr(1) = Format(systemTimeSecondsElapsed(logEntries(0).timestamp, logEntries(i).timestamp), "#.000")
        arr(2) = Format(systemTimeSecondsElapsed(logEntries(i - 1).timestamp, logEntries(i).timestamp), "#.000")
    End If
    
    arr(3) = logEntries(i).description
    
    getOutputforEntry = arr
End Function

' This function converts a datetime in the form of an array of Intergers to a string in ISO format 'yyyy-mm-ddThh:mm:ss.sss'
Private Function timeArrToString(timeArr() As Integer)
    Dim str As String
    str = Format(timeArr(0), "####") & "-" & _
        Format(timeArr(1), "00") & "-" & _
        Format(timeArr(3), "00") & "T" & _
        Format(timeArr(4), "00") & ":" & _
        Format(timeArr(5), "00") & ":" & _
        Format(timeArr(6), "00") & "." & _
        Format(timeArr(7), "000")
    timeArrToString = str
End Function

' This function provides the number of seconds elapsed between two times. The returned number is a double where
' miliseconds are represented by the numbers after the decimal point
Private Function systemTimeSecondsElapsed(from() As Integer, too() As Integer) As Double
    Dim fromMs As Double
    Dim tooMs As Double
    fromMs = DateSerial(from(0), from(1), from(3)) + _
        TimeSerial(from(4), from(5), from(6)) + _
        from(7) / 86400000#
    tooMs = DateSerial(too(0), too(1), too(3)) + _
        TimeSerial(too(4), too(5), too(6)) + _
        too(7) / 86400000#
    systemTimeSecondsElapsed = (tooMs - fromMs) * 24 * 60 * 60
End Function
