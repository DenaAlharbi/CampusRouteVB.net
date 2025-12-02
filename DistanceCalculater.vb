Public Class DistanceCalculator
    Private Const PixelToKmFactor As Double = 0.0005 ' Adjust this based on your map scale

    Public Shared Function CalculateTotalDistance(points As List(Of Point)) As Double
        Dim totalPixels As Double = 0

        For i = 0 To points.Count - 2
            totalPixels += EuclideanDistance(points(i), points(i + 1))
        Next

        Return totalPixels * PixelToKmFactor
    End Function

    Private Shared Function EuclideanDistance(p1 As Point, p2 As Point) As Double
        Dim dx = p2.X - p1.X
        Dim dy = p2.Y - p1.Y
        Return Math.Sqrt(dx * dx + dy * dy)
    End Function
End Class


