Imports System.Data.Entity

Public Class ThingDbInitializer
    Inherits DropCreateDatabaseAlways(Of ThingDbContext)

    Protected Overrides Sub Seed(context As ThingDbContext)
        Dim initialThings As IList(Of Thing) = New List(Of Thing)

        initialThings.Add(New Thing(1, "first - true, positive, all optional values", True, 1, New Date(2016, 1, 1), 1, New Date(2015, 1, 1)))
        initialThings.Add(New Thing(2, "second - false, zero, no optional values", False, 0, New Date(2016, 2, 1)))
        initialThings.Add(New Thing(3, "third - true, negative, optional date", True, -3, New Date(2016, 3, 1), optionalDate:=New Date(2015, 1, 1)))
        initialThings.Add(New Thing(4, "fourth - false, zero, optional integer", False, 0, New Date(2016, 4, 1), 4))
        initialThings.Add(New Thing(5, "fifth - true, positive, zero for optional integer", False, 5, New Date(2016, 4, 1), 0))

        For Each thing As Thing In initialThings
            context.Things.Add(thing)
        Next

        MyBase.Seed(context)
    End Sub

End Class
