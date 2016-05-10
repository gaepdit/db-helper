Imports System.Data.Entity

Public Class ThingDbContext
    Inherits DbContext

    Public Property Things As DbSet(Of Thing)
End Class
