Imports System.Data.Entity

Public Class ThingDbInitializer
    Inherits DropCreateDatabaseAlways(Of ThingDbContext)

    Protected Overrides Sub Seed(context As ThingDbContext)
        For Each thing As Thing In ThingData.ThingsList
            context.Things.Add(thing)
        Next

        MyBase.Seed(context)
    End Sub

End Class
