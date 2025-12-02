Imports Microsoft.Office.Interop
Imports System.Globalization
Public Class ExcelParser
    Private Departments As New Dictionary(Of String, Department)
    Private Instructors As New Dictionary(Of String, Instructor)
    Public CRNMap As New Dictionary(Of Integer, Section)

    Public Sub Parse(filePath As String, Optional maxRows As Integer = Integer.MaxValue)
        Dim xApp As New Excel.Application
        Dim xWorkbook As Excel.Workbook = xApp.Workbooks.Open(filePath)
        Dim xWorksheet As Excel.Worksheet = xWorkbook.Worksheets(1)
        Dim usedRng As Excel.Range = xWorksheet.UsedRange
        Dim numberOfRows As Integer = Math.Min(usedRng.Rows.Count, maxRows)

        For row As Integer = 2 To numberOfRows
            Try
                Dim crnStr = xWorksheet.Cells(row, 2).Value?.ToString()
                If String.IsNullOrWhiteSpace(crnStr) OrElse Not IsNumeric(crnStr) Then Continue For
                Dim crn = Integer.Parse(crnStr)
                Dim term = xWorksheet.Cells(row, 1).Value?.ToString()
                Dim courseCode = xWorksheet.Cells(row, 3).Value?.ToString()
                Dim deptCode = xWorksheet.Cells(row, 4).Value?.ToString()
                If String.IsNullOrWhiteSpace(deptCode) Then deptCode = "GEN"
                Dim sectionNum = xWorksheet.Cells(row, 5).Value?.ToString()
                Dim title = xWorksheet.Cells(row, 6).Value?.ToString()
                Dim activity = xWorksheet.Cells(row, 7).Value?.ToString().Trim()
                If String.IsNullOrWhiteSpace(activity) Then activity = "Unknown"
                Dim daysRaw = xWorksheet.Cells(row, 8).Value?.ToString()
                Dim startTime = xWorksheet.Cells(row, 9).Value?.ToString()
                Dim endTime = xWorksheet.Cells(row, 10).Value?.ToString()
                If Not IsValidTime(startTime) Then startTime = "UNSCHEDULED"
                If Not IsValidTime(endTime) Then endTime = "UNSCHEDULED"
                Dim building = xWorksheet.Cells(row, 11).Value?.ToString()
                If String.IsNullOrWhiteSpace(building) Then building = "UNKNOWN"
                Dim room = xWorksheet.Cells(row, 12).Value?.ToString()
                Dim instructorName = xWorksheet.Cells(row, 13).Value?.ToString()
                If String.IsNullOrWhiteSpace(instructorName) Then instructorName = "Unknown"
                Dim days = ParseDays(daysRaw)

                Dim dept = If(Departments.ContainsKey(deptCode), Departments(deptCode), New Department(deptCode, deptCode))
                Departments(deptCode) = dept
                Dim course = New Course(courseCode, title, dept)
                Dim instructor = If(Instructors.ContainsKey(instructorName), Instructors(instructorName), New Instructor(instructorName))
                Instructors(instructorName) = instructor

                ' Check Duplicate CRNs that are related to each other (ex. LAb+Lec)
                Dim section As Section
                If CRNMap.ContainsKey(crn) Then
                    section = CRNMap(crn)
                Else
                    section = New Section(crn, term, sectionNum, course, instructor)
                    CRNMap(crn) = section
                End If

                Dim slot = New Slot(days, startTime, endTime, building, room, activity)
                section.AddSlot(slot)

            Catch ex As Exception
                Debug.Print($"Row {row} skipped due to error: {ex.Message}")
            End Try
        Next

        xWorkbook.Close(False)
        xApp.Quit()
    End Sub

    Private Function ParseDays(raw As String) As List(Of String)
        Dim result As New List(Of String)
        If String.IsNullOrWhiteSpace(raw) Then Return result
        For Each c As Char In raw.ToUpper()
            If "UMTWR".Contains(c) Then result.Add(c.ToString())
        Next
        Return result
    End Function

    Private Function IsValidTime(t As String) As Boolean
        Dim dt As DateTime
        Return DateTime.TryParseExact(t, "HHmm", CultureInfo.InvariantCulture, DateTimeStyles.None, dt)
    End Function

End Class

