Public Class Instructor
    Public Property Name As String
    Public Property Sections As List(Of Section)

    Public Sub New(name As String)
        Me.Name = name
        Me.Sections = New List(Of Section)()
    End Sub

    Public Sub AddSection(section As Section)
        Sections.Add(section)
    End Sub
End Class

