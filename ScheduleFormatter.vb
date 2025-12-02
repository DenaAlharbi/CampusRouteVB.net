Public Class ScheduleFormatter
    Public Shared Sub AddFormattedSchedule(
        parser As ExcelParser,
        crns As List(Of Integer),
        selectedDayCode As String,
        listBox As ListBox,
        ByRef buildings As HashSet(Of String),
        ByRef courseCount As Integer,
        ByRef index As Integer)

        For Each crn In crns
            If parser.CRNMap.ContainsKey(crn) Then
                Dim section = parser.CRNMap(crn)
                For Each slot In section.Slots
                    If slot.Days.Contains(selectedDayCode) Then
                        listBox.Items.Add($"{index}. {section.Course.Title} ({section.CRN})")
                        listBox.Items.Add($"   Time: {slot.StartTime}–{slot.EndTime}")
                        listBox.Items.Add($"   Location: {slot.Building} {slot.Room}")
                        listBox.Items.Add($"   Instructor: {section.Instructor.Name}")
                        listBox.Items.Add("")
                        buildings.Add(slot.Building)
                        courseCount += 1
                        index += 1
                    End If
                Next
            End If
        Next
    End Sub
    Public Shared Sub AddSummary(listBox As ListBox, selectedDayName As String, courseCount As Integer, buildings As HashSet(Of String))
        listBox.Items.Add($"📚 Total Courses on {selectedDayName}: {courseCount}")
        listBox.Items.Add($"🏛️ Unique Buildings: {buildings.Count}")
    End Sub
    Public Shared Sub AddHeader(listBox As ListBox, selectedDayName As String)
        listBox.Items.Add($"📅 Selected Day: {selectedDayName}")
        listBox.Items.Add("")
    End Sub
    Public Shared Sub AddDistanceSummary(listBox As ListBox, routePoints As List(Of Point))
        Dim distanceKm = DistanceCalculator.CalculateTotalDistance(routePoints)
        listBox.Items.Add($"🧭 Estimated Walking Distance: {Math.Round(distanceKm, 2)} km")
    End Sub
    Public Shared Sub AddCourseListAndPrompt(listBox As ListBox, parser As ExcelParser, crns As List(Of Integer))
        Dim index As Integer = 1
        For Each crn In crns
            If parser.CRNMap.ContainsKey(crn) Then
                Dim section = parser.CRNMap(crn)
                listBox.Items.Add($"{index}. CRN {crn}: {section.Course.Title} - {section.Instructor.Name}")
            Else
                listBox.Items.Add($"{index}. CRN {crn} not found.")
            End If
            index += 1
        Next
        listBox.Items.Add("")
        listBox.Items.Add("📅 Select a day to view your schedule.")
    End Sub

End Class
