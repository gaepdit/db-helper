Imports System.Data.SqlClient
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class ScalarFunctionsTests

    <TestMethod()>
    Public Sub GetSingleValue_String_NoParameters()
        Dim key As Integer = 1
        Dim query As String = "SELECT [Name] from dbo.Things where [ID] = " & key.ToString
        Dim Expected_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Dim Actual_Result As String = DB.GetSingleValue(Of String)(query)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_String_OneParameter()
        Dim key As Integer = 2
        Dim query As String = "SELECT [Name] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Dim Actual_Result As String = DB.GetSingleValue(Of String)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_String_TwoParameters()
        Dim key As Integer = 3
        Dim query As String = "SELECT [Name] from dbo.Things where Status = @status and MandatoryInteger = @mandatoryInteger"
        Dim parameterArray As SqlParameter() = {
            New SqlParameter("@status", ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Status),
            New SqlParameter("@mandatoryInteger", ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).MandatoryInteger)
        }
        Dim Expected_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Dim Actual_Result As String = DB.GetSingleValue(Of String)(query, parameterArray)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_Date()
        Dim key As Integer = 4
        Dim query As String = "SELECT [MandatoryDate] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Date = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).MandatoryDate
        Dim Actual_Result As Date = DB.GetSingleValue(Of Date)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NullableDate_NotNull()
        Dim key As Integer = 1
        Dim query As String = "SELECT [OptionalDate] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Date? = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).OptionalDate
        Dim Actual_Result As Date? = DB.GetSingleValue(Of Date?)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNotNull(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NullableDate_Null()
        Dim key As Integer = 2
        Dim query As String = "SELECT [OptionalDate] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Date? = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).OptionalDate
        Dim Actual_Result As Date? = DB.GetSingleValue(Of Date?)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNull(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NullableInt_NotNull()
        Dim key As Integer = 1
        Dim query As String = "SELECT [OptionalInteger] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Integer? = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).OptionalInteger
        Dim Actual_Result As Integer? = DB.GetSingleValue(Of Integer?)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNotNull(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NullableInt_Null()
        Dim key As Integer = 2
        Dim query As String = "SELECT [OptionalInteger] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Integer? = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).OptionalInteger
        Dim Actual_Result As Integer? = DB.GetSingleValue(Of Integer?)(query, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNull(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_String_FromNull()
        Dim query As String = "SELECT null"
        Dim Expected_Result As String = Nothing
        Dim Actual_Result As String = DB.GetString(query)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNull(Actual_Result)
        Assert.IsTrue(IsNothing(Actual_Result))
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NonNullableInt_FromNull()
        Dim query As String = "SELECT null"
        Dim Expected_Result As Integer = 0
        Dim Actual_Result As Integer = DB.GetSingleValue(Of Integer)(query)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetSingleValue_NonNullableDecimal_FromNull()
        Dim query As String = "SELECT null"
        Dim Expected_Result As Decimal = 0
        Dim Actual_Result As Decimal = DB.GetSingleValue(Of Decimal)(query)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetBoolean_True()
        Dim key As Integer = 1
        Dim query As String = "SELECT [Status] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Actual_Result As Boolean = DB.GetBoolean(query, parameter)
        Assert.IsTrue(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub GetBoolean_False()
        Dim key As Integer = 2
        Dim query As String = "SELECT [Status] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Actual_Result As Boolean = DB.GetBoolean(query, parameter)
        Assert.IsFalse(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub ValueExists_True()
        Dim key As Integer = 1
        Dim query As String = "SELECT [OptionalInteger] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Actual_Result As Boolean = DB.ValueExists(query, parameter)
        Assert.IsTrue(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub ValueExists_False()
        Dim key As Integer = 2
        Dim query As String = "SELECT [OptionalInteger] from dbo.Things where [ID] = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim Actual_Result As Boolean = DB.ValueExists(query, parameter)
        Assert.IsFalse(Actual_Result)
    End Sub

End Class