Public Class CBStopWatch
    Private priStopWatch As New Stopwatch
    Private arrEvents() As EventInfo
    Private _Name As String

    Structure EventInfo
        Public Start_Time As Double
        Public Total_Run_Time As Double
        Public Time_Since_Last_Event As Double
        Public Event_Description As String
    End Structure
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set
            _Name = Value
        End Set
    End Property


    Public Sub Start_Stopwatch()
        Dim newEvent As New EventInfo
        newEvent.Start_Time = 0
        newEvent.Event_Description = "Timer Started"
        newEvent.Time_Since_Last_Event = 0

        ReDim arrEvents(0)
        arrEvents(0) = newEvent
        priStopWatch.Start()
    End Sub

    Public Sub Resume_Stopwatch()
        priStopWatch.Start()
        Dim newEvent As New EventInfo
        newEvent.Start_Time = priStopWatch.ElapsedMilliseconds
        newEvent.Event_Description = "Timer Resumed"

        Dim arrSize As Integer = arrEvents.Length
        newEvent.Time_Since_Last_Event = newEvent.Start_Time - arrEvents(arrSize - 1).Start_Time
        newEvent.Total_Run_Time = newEvent.Start_Time - arrEvents(0).Start_Time


        ReDim Preserve arrEvents(arrSize)
        arrEvents(arrSize) = newEvent
    End Sub


    Public Sub Pause_Stopwatch()
        priStopWatch.Stop()
        Dim newEvent As New EventInfo
        newEvent.Start_Time = priStopWatch.ElapsedMilliseconds
        newEvent.Event_Description = "Timer Paused"

        Dim arrSize As Integer = arrEvents.Length
        newEvent.Time_Since_Last_Event = newEvent.Start_Time - arrEvents(arrSize - 1).Start_Time
        newEvent.Total_Run_Time = newEvent.Start_Time - arrEvents(0).Start_Time


        ReDim Preserve arrEvents(arrSize)
        arrEvents(arrSize) = newEvent
    End Sub
    Public Sub Stop_Stopwatch()
        priStopWatch.Stop()
        Dim newEvent As New EventInfo
        newEvent.Start_Time = priStopWatch.ElapsedMilliseconds
        newEvent.Event_Description = "Timer Stopped"

        Dim arrSize As Integer = arrEvents.Length
        newEvent.Time_Since_Last_Event = newEvent.Start_Time - arrEvents(arrSize - 1).Start_Time
        newEvent.Total_Run_Time = newEvent.Start_Time - arrEvents(0).Start_Time


        ReDim Preserve arrEvents(arrSize)
        arrEvents(arrSize) = newEvent

    End Sub
    Public Sub Reset_Stopwatch()
        'Clear the array of events
        If Not arrEvents Is Nothing Then
            Array.Clear(arrEvents, 0, arrEvents.Length)
        End If
        priStopWatch.Reset()
    End Sub
    Public Sub RecordEvent(Event_Description As String)
        Dim newEvent As New EventInfo
        newEvent.Start_Time = priStopWatch.ElapsedMilliseconds

        Dim arrSize As Integer = arrEvents.Length
        newEvent.Time_Since_Last_Event = newEvent.Start_Time - arrEvents(arrSize - 1).Start_Time
        newEvent.Event_Description = Event_Description
        newEvent.Total_Run_Time = newEvent.Start_Time - arrEvents(0).Start_Time
        ReDim Preserve arrEvents(arrSize)
        arrEvents(arrSize) = newEvent


    End Sub

    Public Sub ListEvents()
        Console.Write(vbCrLf & _Name & " EVENTS LIST" & vbCrLf)
        For Each e As EventInfo In arrEvents
            Console.WriteLine(e.Start_Time.ToString & "ms - " & e.Time_Since_Last_Event & "ms - " & e.Total_Run_Time & "ms - " & e.Event_Description)
        Next
    End Sub
End Class
