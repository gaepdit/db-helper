Imports System.Data.SqlClient

<TestClass()>
Public Class NonQueryFunctionsTests

    <TestMethod()>
    Public Sub Insert_NoParameters()
        Dim key As Integer = 13
        Dim name As String = "test" & key.ToString
        Dim rowsAffected As Integer = 0
        Dim query As String = "INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (N'" & name & "', 0, " & key & ", N'2016-04-01 00:00:00')"
        Dim Actual_Result As Boolean = DB.RunCommand(query, rowsAffected:=rowsAffected)
        Assert.IsTrue(Actual_Result)
        Assert.AreEqual(1, rowsAffected)

        Dim extraQuery As String = "SELECT MandatoryInteger FROM dbo.Things WHERE Name = @name"
        Dim extraParameter As New SqlParameter("@name", name)
        Dim Extra_Result As Integer = DB.GetSingleValue(Of Integer)(extraQuery, extraParameter)
        Assert.AreEqual(key, Extra_Result)
    End Sub

    <TestMethod()>
    Public Sub Insert_WithOneParameter()
        Dim key As Integer = 14
        Dim name As String = "test" & key.ToString
        Dim rowsAffected As Integer = 0
        Dim query As String = "INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (N'" & name & "', 0, @status , N'2016-04-01 00:00:00')"
        Dim parameter As New SqlParameter("@status", key)
        Dim Actual_Result As Boolean = DB.RunCommand(query, parameter, rowsAffected)
        Assert.IsTrue(Actual_Result)
        Assert.AreEqual(1, rowsAffected)

        Dim extraQuery As String = "SELECT MandatoryInteger FROM dbo.Things WHERE Name = @name"
        Dim extraParameter As New SqlParameter("@name", name)
        Dim Extra_Result As Integer = DB.GetSingleValue(Of Integer)(extraQuery, extraParameter)
        Assert.AreEqual(key, Extra_Result)
    End Sub

    <TestMethod()>
    Public Sub Insert_WithMultipleParameter()
        Dim key As Integer = 15
        Dim insertName As String = "test" & key.ToString
        Dim insertDate As Date = New Date(2016, 2, 3)

        Dim rowsAffected As Integer = 0
        Dim query As String = "INSERT INTO dbo.Things (Name, Status, MandatoryInteger, MandatoryDate) VALUES (@name, @status, @mandatoryInteger, @mandatoryDate)"
        Dim parameterArray() As SqlParameter = {
            New SqlParameter("@name", insertName),
            New SqlParameter("@status", True),
            New SqlParameter("@mandatoryInteger", key),
            New SqlParameter("@mandatoryDate", insertDate)
        }
        Dim Actual_Result As Boolean = DB.RunCommand(query, parameterArray, rowsAffected)
        Assert.IsTrue(Actual_Result)
        Assert.AreEqual(1, rowsAffected)

        Dim extraQuery1 As String = "SELECT Status FROM dbo.Things WHERE Name = @name"
        Dim extraParameter1 As New SqlParameter("@name", insertName)
        Dim Extra_Result1 As Boolean = DB.GetBoolean(extraQuery1, extraParameter1)
        Assert.IsTrue(Extra_Result1)

        Dim extraQuery2 As String = "SELECT MandatoryInteger FROM dbo.Things WHERE Name = @name"
        Dim extraParameter2 As New SqlParameter("@name", insertName)
        Dim Extra_Result2 As Integer = DB.GetSingleValue(Of Integer)(extraQuery2, extraParameter2)
        Assert.AreEqual(key, Extra_Result2)

        Dim extraQuery3 As String = "SELECT MandatoryDate FROM dbo.Things WHERE Name = @name"
        Dim extraParameter3 As New SqlParameter("@name", insertName)
        Dim Extra_Result3 As Date = DB.GetSingleValue(Of Date)(extraQuery3, extraParameter3)
        Assert.AreEqual(insertDate, Extra_Result3)

        Dim extraQuery4 As String = "SELECT OptionalInteger FROM dbo.Things WHERE Name = @name"
        Dim extraParameter4 As New SqlParameter("@name", insertName)
        Dim Extra_Result4 As Integer? = DB.GetSingleValue(Of Integer?)(extraQuery4, extraParameter4)
        Assert.IsNull(Extra_Result4)

        Dim extraQuery5 As String = "SELECT OptionalDate FROM dbo.Things WHERE Name = @name"
        Dim extraParameter5 As New SqlParameter("@name", insertName)
        Dim Extra_Result5 As Date? = DB.GetSingleValue(Of Date?)(extraQuery5, extraParameter5)
        Assert.IsNull(Extra_Result5)
    End Sub

End Class