Imports System.Data.Entity
Imports EpdIt

Module CommonDbObjects

    Public Property ConnectionString As String
    Public Property DB As DBHelper

End Module

<TestClass>
Public Class Setup

#Disable Warning IDE0060 ' Remove unused parameter
    <AssemblyInitialize>
    Public Shared Sub DbSetup(testContext As TestContext)
#Enable Warning IDE0060 ' Remove unused parameter
        ConnectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=EpdItDbHelperTests.ThingDbContext;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False"
        DB = New DBHelper(ConnectionString)
        InitializeTestDb()
    End Sub

    Private Shared Sub InitializeTestDb()
        Database.SetInitializer(New ThingDbInitializer)

        Dim context As New ThingDbContext
        context.Database.Initialize(True)

        For Each dbo As String In DbTestObjectStrings
            context.Database.ExecuteSqlCommand(dbo)
        Next
    End Sub

    <AssemblyCleanup>
    Public Shared Sub DbTeardown()
        DB = Nothing
    End Sub

End Class
