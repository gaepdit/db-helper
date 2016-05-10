Imports EpdItDbHelper
Imports System.Data.SqlClient

Namespace EpdItDbHelper.Tests

    <TestClass()>
    Public Class ScalarFunctionsTests

        <TestMethod()>
        Public Sub GetSingleValue_NoParameters()
            Dim query As String = "SELECT DATEDIFF(day,'2014-06-05','2014-08-05') AS DiffDate"
            Dim Expected_Date_Diff As Integer = 61
            Dim result As Integer = DBInstance.GetSingleValue(Of Integer)(query)
            Assert.AreEqual(Expected_Date_Diff, result)
        End Sub

        <TestMethod()>
        Public Sub GetSingleValue_OneParameter()
            Dim query As String = "SELECT DATEDIFF(day, @day1,'2014-08-05') AS DiffDate"
            Dim parameter As New SqlParameter("day1", "2014-06-05")
            Dim Expected_Date_Diff As Integer = 61
            Dim result As Integer = DBInstance.GetSingleValue(Of Integer)(query, parameter)
            Assert.AreEqual(Expected_Date_Diff, result)
        End Sub

        <TestMethod()>
        Public Sub GetSingleValue_TwoParameters()
            Dim query As String = "SELECT DATEDIFF(day, @day1,@day2) AS DiffDate"
            Dim parameterArray As SqlParameter() = {
            New SqlParameter("day1", "2014-06-05"),
            New SqlParameter("day2", "2014-08-05")
        }
            Dim Expected_Date_Diff As Integer = 61
            Dim result As Integer = DBInstance.GetSingleValue(Of Integer)(query, parameterArray)
            Assert.AreEqual(Expected_Date_Diff, result)
        End Sub

        <TestMethod()>
        Public Sub ValueExists_NoParameters()
            Dim query As String = "select getdate()"
            Dim result As String = DBInstance.ValueExists(query)
            Assert.IsTrue(result)
        End Sub

        <TestMethod()>
        Public Sub ValueExists_OneParameter()
            Dim query As String = "SELECT DATEDIFF(day, @day1,'2014-08-05') AS DiffDate"
            Dim parameter As New SqlParameter("day1", "2014-06-05")
            Dim result As String = DBInstance.ValueExists(query, parameter)
            Assert.IsTrue(result)
        End Sub

        <TestMethod()>
        Public Sub ValueExists_TwoParameters()
            Dim query As String = "SELECT DATEDIFF(day, @day1,@day2) AS DiffDate"
            Dim parameterArray As SqlParameter() = {
            New SqlParameter("day1", "2014-06-05"),
            New SqlParameter("day2", "2014-08-05")
        }
            Dim result As String = DBInstance.ValueExists(query, parameterArray)
            Assert.IsTrue(result)
        End Sub

    End Class

End Namespace