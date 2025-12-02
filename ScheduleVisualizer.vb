Public Class ScheduleVisualizer
    Public Shared Sub DrawRoute(g As Graphics, routePoints As List(Of Point))
        If routePoints Is Nothing OrElse routePoints.Count = 0 Then Exit Sub

        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim font As New Font("Arial", 16, FontStyle.Bold)
        Dim textBrush As New SolidBrush(Color.White)
        Dim bgBrush As New SolidBrush(Color.Black)
        Dim pen As New Pen(Color.Black, 6)
        pen.EndCap = Drawing2D.LineCap.ArrowAnchor

        For i = 0 To routePoints.Count - 1
            Dim pt = routePoints(i)
            If i < routePoints.Count - 1 Then
                g.DrawLine(pen, pt, routePoints(i + 1))
            End If
            Dim label = (i + 1).ToString()
            Dim labelSize = g.MeasureString(label, font)
            Dim offsetX = 5 + (i Mod 3) * 10
            Dim offsetY = -20 - (i Mod 2) * 10
            Dim circleRadius = 35
            Dim centerX = pt.X
            Dim centerY = pt.Y

            ' Draw black circle
            g.DrawEllipse(pen, centerX - circleRadius \ 2, centerY - circleRadius \ 2, circleRadius, circleRadius)
            ' Center the label inside the circle
            Dim labelX = centerX - labelSize.Width \ 2
            Dim labelY = centerY - labelSize.Height \ 2
            g.DrawString(label, font, Brushes.White, labelX, labelY)


        Next
    End Sub
End Class
