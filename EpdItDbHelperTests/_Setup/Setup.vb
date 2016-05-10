Imports System.Data.Entity
Imports EpdIt

Module CommonDbObjects

    Public Property ConnectionString As String
    Public Property DB As DBHelper

End Module

<TestClass>
Public Class Setup

    <AssemblyInitialize>
    Public Shared Sub DbSetup(testContext As TestContext)
        ConnectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=EpdItDbHelperTests.ThingDbContext;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"
        DB = New DBHelper(ConnectionString)
        InitializeTestDb()
    End Sub

    Private Shared Sub InitializeTestDb()
        Database.SetInitializer(New ThingDbInitializer)
        Dim context As New ThingDbContext
        context.Database.Initialize(True)
    End Sub

    <AssemblyCleanup>
    Public Shared Sub DbTeardown()
        DB = Nothing
    End Sub

End Class
