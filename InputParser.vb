Public Class InputParser
    ' Parses a string into a list of integers, filtering out invalid entries
    Public Shared Function ParseIntegerList(input As String, Optional ByRef invalidEntries As List(Of String) = Nothing) As List(Of Integer)
        Dim result As New List(Of Integer)
        invalidEntries = New List(Of String)

        Dim tokens = input.Split({" "c, ","c, vbTab, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each token In tokens
            Dim trimmed = token.Trim()
            If Integer.TryParse(trimmed, Nothing) Then
                result.Add(Integer.Parse(trimmed))
            Else
                invalidEntries.Add(trimmed)
            End If
        Next

        Return result
    End Function

End Class

