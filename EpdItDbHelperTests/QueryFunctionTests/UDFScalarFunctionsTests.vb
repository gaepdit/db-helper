Imports System.Data.SqlClient

<TestClass()>
Public Class UDFScalarFunctionsTests

    <TestMethod()>
    Public Sub UDF_NoParameters()
        Dim udfQuery As String = "select dbo." & DBO_UDF_NoParameters_NAME & "()"
        Dim Actual_Result As Integer = DB.GetSingleValue(Of Integer)(udfQuery)
        Assert.AreEqual(DBO_UDF_NoParameters_VALUE, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub UDF_WithParameters()
        Dim p1 As Integer = 3
        Dim p2 As Integer = 10
        Dim udfQuery As String = "select dbo." & DBO_UDF_WithParameters_NAME & "(@param1, @param2)"
        Dim parameterArray As SqlParameter() = {
            New SqlParameter("param1", p1),
            New SqlParameter("param2", p2)
        }
        Dim Expected_Result As Integer = p1 + p2
        Dim Actual_Result As Integer = DB.GetSingleValue(Of Integer)(udfQuery, parameterArray)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

End Class