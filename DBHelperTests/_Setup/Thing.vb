Public Class Thing

    Public Property ID As Integer
    Public Property Name As String
    Public Property Status As Boolean
    Public Property MandatoryInteger As Integer
    Public Property MandatoryDate As Date
    Public Property OptionalInteger As Integer?
    Public Property OptionalDate As Date?

    Public Sub New(id As Integer, name As String, status As Boolean, mandatoryInteger As Integer, mandatoryDate As Date, Optional optionalInteger As Integer? = Nothing, Optional optionalDate As Date? = Nothing)
        Me.ID = id
        Me.Name = name
        Me.Status = status
        Me.MandatoryInteger = mandatoryInteger
        Me.MandatoryDate = mandatoryDate
        Me.OptionalInteger = optionalInteger
        Me.OptionalDate = optionalDate
    End Sub

End Class
