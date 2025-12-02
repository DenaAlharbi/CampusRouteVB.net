Public Class Slot
    Public Property Days As List(Of String)
    Public Property StartTime As String
    Public Property EndTime As String
    Public Property Building As String
    Public Property Room As String
    Public Property ActivityType As String

    Public Sub New(days As List(Of String), startTime As String, endTime As String,
                   building As String, room As String, activityType As String)
        Me.Days = days
        Me.StartTime = startTime
        Me.EndTime = endTime
        Me.Building = building
        Me.Room = room
        Me.ActivityType = activityType
    End Sub

    Public Overrides Function ToString() As String
        Return $"{ActivityType} | {Building} {Room} | {StartTime} - {EndTime} | Days: {String.Join(",", Days)}"
    End Function
End Class

