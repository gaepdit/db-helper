Imports System.Data.SqlClient

<TestClass()>
Public Class UDFScalarFunctionsTests

    <TestMethod()>
    Public Sub CallFunction_Integer_NoParameters()
        Dim udfQuery As String = "select dbo." & DB_TEST_FUNCTION_A & "()"
        Dim Expected_Result As Integer = DB_TEST_INT_A
        Dim Actual_Result As Integer = DB.GetSingleValue(Of Integer)(udfQuery)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub CallFunction_Integer_WithParameters()
        Dim p1 As Integer = 3
        Dim p2 As Integer = 10
        Dim udfQuery As String = "select dbo." & DB_TEST_FUNCTION_B & "(@param1, @param2)"
        Dim parameterArray As SqlParameter() = {
            New SqlParameter("param1", p1),
            New SqlParameter("param2", p2)
        }
        Dim Expected_Result As Integer = p1 + p2
        Dim Actual_Result As Integer = DB.GetSingleValue(Of Integer)(udfQuery, parameterArray)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

End Class