Imports System.Data.Entity

Public Class ThingDbInitializer
    Inherits DropCreateDatabaseAlways(Of ThingDbContext)

    Protected Overrides Sub Seed(context As ThingDbContext)
        If context.Database.Exists Then
            context.Database.Delete()
        End If

        context.Database.Create()

        For Each thing As Thing In ThingData.Things
            context.Things.Add(thing)
        Next

        MyBase.Seed(context)
    End Sub

End Class
