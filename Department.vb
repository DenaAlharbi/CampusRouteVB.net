Public Class Department
    Private Property DeptCode As String
    Private Property Name As String
    Public Property Courses As List(Of Course)

    Public Sub New(deptCode As String, name As String)
        Me.DeptCode = deptCode
        Me.Name = name
        Me.Courses = New List(Of Course)()
    End Sub

    Public Sub AddCourse(course As Course)
        Courses.Add(course)
    End Sub
End Class

