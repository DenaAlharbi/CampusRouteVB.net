Public Class Course
    Private Property CourseCode As String
    Public Property Title As String
    Private Property Department As Department
    Public Property Sections As List(Of Section)

    Public Sub New(courseCode As String, title As String, department As Department)
        Me.CourseCode = courseCode
        Me.Title = title
        Me.Department = department
        Me.Sections = New List(Of Section)()
        department.AddCourse(Me)
    End Sub

    Public Sub AddSection(section As Section)
        Sections.Add(section)
    End Sub
End Class

