Imports System.Data.SqlClient

<TestClass()>
Public Class SPNonQueryFunctionsTests

    <TestMethod()>
    Public Sub SPInsert_NoParameters()
        Dim rowsAffected As Integer = 0
        Dim spName As String = DB_TEST_PROCEDURE_C
        Dim Actual_Result As Boolean = DB.SPRunCommand(spName, rowsAffected:=rowsAffected)
        Assert.IsTrue(Actual_Result)
        Assert.AreEqual(1, rowsAffected)

        Dim extraQuery As String = "SELECT MandatoryInteger FROM dbo.Things WHERE Name = @name"
        Dim extraParameter As New SqlParameter("@name", DB_TEST_STRING_C)
        Dim Extra_Result As Integer = DB.GetSingleValue(Of Integer)(extraQuery, extraParameter)
        Assert.AreEqual(DB_TEST_INT_C, Extra_Result)
    End Sub

    <TestMethod()>
    Public Sub SPInsert_WithMultipleParameter()
        Dim key As Integer = 21
        Dim insertName As String = "test" & key.ToString
        Dim insertDate As Date = New Date(2016, 2, 3)

        Dim spName As String = DB_TEST_PROCEDURE_D
        Dim parameterArray() As SqlParameter = {
            New SqlParameter("@name", insertName),
            New SqlParameter("@status", True),
            New SqlParameter("@mandatoryInteger", key),
            New SqlParameter("@mandatoryDate", insertDate)
        }
        Dim Actual_Result As Boolean = DB.SPRunCommand(spName, parameterArray)
        Assert.IsTrue(Actual_Result)

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