Public Class Section
    Public Property CRN As Integer
    Public Property Term As String
    Public Property SectionNumber As String
    Public Property Course As Course
    Public Property Instructor As Instructor
    Public Property Slots As List(Of Slot)

    Public Sub New(crn As Integer, term As String, sectionNumber As String, course As Course, instructor As Instructor)
        Me.CRN = crn
        Me.Term = term
        Me.SectionNumber = sectionNumber
        Me.Course = course
        Me.Instructor = instructor
        Me.Slots = New List(Of Slot)()
        course.AddSection(Me)
        instructor.AddSection(Me)
    End Sub

    Public Sub AddSlot(slot As Slot)
        Slots.Add(slot)
    End Sub
End Class


