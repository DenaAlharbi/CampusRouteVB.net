Imports System.Globalization
Imports Microsoft.Office.Interop

Public Class Form1
    Private lastSelectedDayButton As Button = Nothing


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AddHandler MapPictureBox.Paint, AddressOf MapPictureBox_Paint
        AddHandler MapPictureBox.MouseClick, AddressOf MapPictureBox_MouseClick
        AddHandler xListBox.SelectedIndexChanged, AddressOf xListBox_SelectedIndexChanged


    End Sub

    Private parser As ExcelParser
    Private Sub ReadSheetBtn_Click(sender As Object, e As EventArgs) Handles ReadSheetBtn.Click
        Try
            StatusLabel.Text = "📄 Parsing Excel file..."
            StatusLabel.Visible = True
            StatusLabel.Refresh() ' Force UI update

            Dim timeStart As DateTime = Now
            parser = New ExcelParser()
            parser.Parse("C:\Users\denaa\OneDrive\Term251-dena.xlsx", )

            xListBox.Items.Clear()
            xListBox.Items.Add("Parsed CRNs: " & parser.CRNMap.Count.ToString())

            For Each kvp In parser.CRNMap
                Dim section = kvp.Value
                xListBox.Items.Add($"CRN {section.CRN}: {section.Course.Title} - {section.Instructor.Name}")
            Next
            StatusLabel.Text = "✅ Schedule ready."
            Log("Total Time= " & CStr(CInt((Now - timeStart).TotalMilliseconds)))
        Catch ex As Exception
            StatusLabel.Text = "❌ Error loading file."
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub xListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim selectedText As String = xListBox.SelectedItem?.ToString()
        If String.IsNullOrEmpty(selectedText) Then Exit Sub

        ' Try to extract CRN from the line
        Dim match = System.Text.RegularExpressions.Regex.Match(selectedText, "CRN\s+(\d+)")
        If match.Success Then
            Dim crn = match.Groups(1).Value
            Clipboard.SetText(crn)
            MsgBox($"CRN {crn} copied to clipboard.")
        End If
    End Sub


    Private Sub DrawBtn_Click(sender As Object, e As EventArgs) Handles DrawBtn.Click

        If parser Is Nothing Then
            MsgBox("Please read the Excel file first.")
            Return
        End If
        ScheduleListBox.Items.Clear()


        Dim invalids As New List(Of String)
        Dim inputCRNs = InputParser.ParseIntegerList(CRNTextBox.Text, invalids)
        If inputCRNs.Count = 0 Then
            ScheduleListBox.Items.Add("⚠️ No valid CRNs entered.")
            Return
        End If

        If invalids.Count > 0 Then
            For Each item In invalids
                ScheduleListBox.Items.Add($"⚠️ Invalid CRN: {item}")
            Next
        End If
        If String.IsNullOrEmpty(selectedDay) Then
            ScheduleFormatter.AddCourseListAndPrompt(ScheduleListBox, parser, inputCRNs)
            Return
        End If
    End Sub

    Private selectedDay As String = ""

    Private Sub DayButton_Click(sender As Object, e As EventArgs) Handles BtnU.Click, BtnM.Click, BtnT.Click, BtnW.Click, BtnR.Click
        If parser Is Nothing Then
            MsgBox("Please read the Excel file first.")
            Return
        End If
        Dim clickedBtn = CType(sender, Button)
        Dim selectedDayCode = clickedBtn.Tag.ToString()
        selectedDay = selectedDayCode

        ' 🔄 Reset previous button style
        If lastSelectedDayButton IsNot Nothing Then
            lastSelectedDayButton.BackColor = SystemColors.Control
            lastSelectedDayButton.Font = New Font(lastSelectedDayButton.Font, FontStyle.Regular)
        End If

        ' ✅ Highlight current button
        clickedBtn.BackColor = Color.LightBlue
        clickedBtn.Font = New Font(clickedBtn.Font, FontStyle.Bold)
        lastSelectedDayButton = clickedBtn

        Dim selectedDayName = ScheduleManager.GetDayName(selectedDayCode)

        ' Get CRNs from textbox
        Dim inputCRNs = InputParser.ParseIntegerList(CRNTextBox.Text)
        ScheduleListBox.Items.Clear()
        ScheduleFormatter.AddHeader(ScheduleListBox, selectedDayName)
        routePoints.Clear()

        Dim orderedSlots = ScheduleManager.GetSlotsForDay(parser, inputCRNs, selectedDayCode)
        For Each slot In orderedSlots
            If buildingLocations.ContainsKey(slot.Building) Then
                routePoints.Add(buildingLocations(slot.Building))
            End If
        Next
        MapPictureBox.Invalidate() ' Trigger redraw
        Dim index = 1
        Dim buildings As New HashSet(Of String)
        Dim courseCount = 0

        ScheduleFormatter.AddFormattedSchedule(parser, inputCRNs, selectedDayCode, ScheduleListBox, buildings, courseCount, index)
        ScheduleFormatter.AddDistanceSummary(ScheduleListBox, routePoints)
        ScheduleFormatter.AddSummary(ScheduleListBox, selectedDayName, courseCount, buildings)
    End Sub

    Private routePoints As New List(Of Point)
    Private Sub MapPictureBox_Paint(sender As Object, e As PaintEventArgs)
        ScheduleVisualizer.DrawRoute(e.Graphics, routePoints)
    End Sub

    Private buildingLocations As New Dictionary(Of String, Point) From {
    {"22", New Point(737, 617)},
    {"59", New Point(652, 359)},
    {"23", New Point(793, 664)},
    {"24", New Point(768, 738)},
    {"25", New Point(816, 789)},
    {"11", New Point(692, 729)},
    {"26", New Point(191, 123)},
    {"15", New Point(280, 419)},
    {"10", New Point(610, 622)},
    {"18", New Point(482, 633)},
    {"7", New Point(506, 432)},
    {"6", New Point(508, 375)},
    {"4", New Point(374, 228)},
    {"14", New Point(430, 547)},
    {"40", New Point(605, 57)}}
    ' Add more buildings as needed
    Private Sub MapPictureBox_MouseClick(sender As Object, e As MouseEventArgs)
        Dim clickedPoint As Point = e.Location
        MsgBox($"Clicked at: X={clickedPoint.X}, Y={clickedPoint.Y}")

    End Sub

    Public Shared Sub Log(str As String)
        Try
            System.Diagnostics.Debug.Print(str)
        Catch ex As Exception
            'MsgBox("Error " + ex.StackTrace())
        End Try
    End Sub

End Class
