Public Class ThingData

    Private Sub New()
    End Sub

    Public Shared Property ThingsList As New List(Of Thing) From {
        New Thing(1, "first - true, positive, all optional values", True, 1, New Date(2016, 1, 1), 1, New Date(2015, 1, 1)),
        New Thing(2, "second - false, zero, no optional values", False, 0, New Date(2016, 2, 1)),
        New Thing(3, "third - true, negative, only optional date", True, -3, New Date(2016, 3, 1), optionalDate:=New Date(2015, 1, 1)),
        New Thing(4, "fourth - true, zero, only optional integer", True, 0, New Date(2016, 4, 1), 4),
        New Thing(5, "fifth - true, positive, zero for optional integer", False, 5, New Date(2016, 4, 1), 0)
    }

End Class
