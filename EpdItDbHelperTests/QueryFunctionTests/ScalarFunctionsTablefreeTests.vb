Imports System.Data.SqlClient

Namespace EpdItDbHelper.Tests

    <TestClass()>
    Public Class ScalarFunctionsTablefreeTests

        <TestMethod()>
        Public Sub TF_GetSingleValue_NoParameters()
            Dim query As String = "SELECT DATEDIFF(day,'2014-06-05','2014-08-05') AS DiffDate"
            Dim Expected_Result As Integer = 61
            Dim Actual_Result As Integer = DBInstance.GetSingleValue(Of Integer)(query)
            Assert.AreEqual(Expected_Result, Actual_Result)
        End Sub

        <TestMethod()>
        Public Sub TF_GetSingleValue_OneParameter()
            Dim query As String = "SELECT DATEDIFF(day, @day1,'2014-08-05') AS DiffDate"
            Dim parameter As New SqlParameter("day1", "2014-06-05")
            Dim Expected_Result As Integer = 61
            Dim Actual_Result As Integer = DBInstance.GetSingleValue(Of Integer)(query, parameter)
            Assert.AreEqual(Expected_Result, Actual_Result)
        End Sub

        <TestMethod()>
        Public Sub TF_GetSingleValue_TwoParameters()
            Dim query As String = "SELECT DATEDIFF(day, @day1,@day2) AS DiffDate"
            Dim parameterArray As SqlParameter() = {
                New SqlParameter("day1", "2014-06-05"),
                New SqlParameter("day2", "2014-08-05")
            }
            Dim Expected_Result As Integer = 61
            Dim Actual_Result As Integer = DBInstance.GetSingleValue(Of Integer)(query, parameterArray)
            Assert.AreEqual(Expected_Result, Actual_Result)
        End Sub

        <TestMethod()>
        Public Sub TF_ValueExists_NoParameters()
            Dim query As String = "select getdate()"
            Dim Actual_Result As Boolean = DBInstance.ValueExists(query)
            Assert.IsTrue(Actual_Result)
        End Sub

        <TestMethod()>
        Public Sub TF_ValueExists_OneParameter()
            Dim query As String = "SELECT DATEDIFF(day, @day1,'2014-08-05') AS DiffDate"
            Dim parameter As New SqlParameter("day1", "2014-06-05")
            Dim Actual_Result As Boolean = DBInstance.ValueExists(query, parameter)
            Assert.IsTrue(Actual_Result)
        End Sub

        <TestMethod()>
        Public Sub TF_ValueExists_TwoParameters()
            Dim query As String = "SELECT DATEDIFF(day, @day1,@day2) AS DiffDate"
            Dim parameterArray As SqlParameter() = {
                New SqlParameter("day1", "2014-06-05"),
                New SqlParameter("day2", "2014-08-05")
            }
            Dim Actual_Result As Boolean = DBInstance.ValueExists(query, parameterArray)
            Assert.IsTrue(Actual_Result)
        End Sub

    End Class

End Namespace