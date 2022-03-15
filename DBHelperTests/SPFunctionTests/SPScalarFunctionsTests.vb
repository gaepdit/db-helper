Imports System.Data.SqlClient
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class SPScalarFunctionsTests

    <TestMethod()>
    Public Sub SPGetBoolean_True()
        Dim spName As String = "SpGetBoolean"
        Dim parameter As New SqlParameter("@select", True)
        Assert.IsTrue(DB.SPGetBoolean(spName, parameter))
    End Sub

    <TestMethod()>
    Public Sub SPGetBoolean_False()
        Dim spName As String = "SpGetBoolean"
        Dim parameter As New SqlParameter("@select", False)
        Assert.IsFalse(DB.SPGetBoolean(spName, parameter))
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_String_Null()
        Dim spName As String = "SpGetString"
        Dim parameter As New SqlParameter("@select", 0)
        Dim result As String = DB.SPGetSingleValue(Of String)(spName, parameter)
        Assert.IsNull(result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_String_Empty()
        Dim spName As String = "SpGetString"
        Dim parameter As New SqlParameter("@select", 1)
        Dim result As String = DB.SPGetSingleValue(Of String)(spName, parameter)
        Assert.IsNotNull(result)
        Assert.AreEqual("", result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_String_Nonempty()
        Dim spName As String = "SpGetString"
        Dim parameter As New SqlParameter("@select", 2)
        Dim result As String = DB.SPGetSingleValue(Of String)(spName, parameter)
        Assert.AreEqual("abc", result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_Int_Zero()
        Dim spName As String = "SpGetInt"
        Dim parameter As New SqlParameter("@select", 0)
        Dim result As Integer = DB.SPGetSingleValue(Of Integer)(spName, parameter)
        Assert.AreEqual(0, result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_Int_Pos()
        Dim spName As String = "SpGetInt"
        Dim parameter As New SqlParameter("@select", 1)
        Dim result As Integer = DB.SPGetSingleValue(Of Integer)(spName, parameter)
        Assert.AreEqual(1, result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_Int_Net()
        Dim spName As String = "SpGetInt"
        Dim parameter As New SqlParameter("@select", 2)
        Dim result As Integer = DB.SPGetSingleValue(Of Integer)(spName, parameter)
        Assert.AreEqual(-1, result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_NullableInt_Value()
        Dim spName As String = "SpGetNullableInt"
        Dim parameter As New SqlParameter("@select", 1)
        Dim result As Integer? = DB.SPGetSingleValue(Of Integer?)(spName, parameter)
        Assert.AreEqual(1, result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_NullableInt_Null()
        Dim spName As String = "SpGetNullableInt"
        Dim parameter As New SqlParameter("@select", 0)
        Dim result As Integer? = DB.SPGetSingleValue(Of Integer?)(spName, parameter)
        Assert.IsNull(result)
        Assert.AreEqual(Nothing, result)
        Assert.AreEqual(0, EpdIt.DBUtilities.GetNullable(Of Integer)(result))
    End Sub

    <TestMethod()>
    Public Sub SPGetBooleanAndReturnValue_True()
        Dim spName As String = "SpGetBooleanAndReturn"
        Dim parameter As New SqlParameter("@select", True)
        Dim returnValue As Integer = 0
        Dim result As Boolean = DB.SPGetBoolean(spName, parameter, returnValue:=returnValue)
        Assert.IsTrue(result)
        Assert.AreEqual(2, returnValue)
    End Sub

End Class