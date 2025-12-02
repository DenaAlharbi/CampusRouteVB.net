Imports System.Globalization

Public Class ScheduleManager
    Public Shared Function GetSlotsForDay(parser As ExcelParser, crns As List(Of Integer), dayCode As String) As List(Of Slot)
        Dim slots As New List(Of Slot)
        For Each crn In crns
            If parser.CRNMap.ContainsKey(crn) Then
                Dim section = parser.CRNMap(crn)
                For Each slot In section.Slots
                    If slot.Days.Contains(dayCode) Then
                        slots.Add(slot)
                    End If
                Next
            End If
        Next
        Return slots.OrderBy(Function(s)
                                 Dim dt As DateTime
                                 If DateTime.TryParseExact(s.StartTime, "HHmm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
                                     Return dt
                                 Else
                                     Return DateTime.MaxValue ' Push unscheduled to end
                                 End If
                             End Function).ToList()
    End Function
    Public Shared Function GetDayName(code As String) As String
        Dim dayNames As New Dictionary(Of String, String) From {
            {"U", "Sunday"},
            {"M", "Monday"},
            {"T", "Tuesday"},
            {"W", "Wednesday"},
            {"R", "Thursday"}
        }
        Return If(dayNames.ContainsKey(code), dayNames(code), "Unknown")
    End Function


End Class

