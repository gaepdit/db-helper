Imports System.Data.SqlClient

<TestClass()>
Public Class SPScalarFunctionsTests

    <TestMethod()>
    Public Sub SPGetSingleValue_String()
        Dim key As Integer = DBO_SP_Return_String_VALUE
        Dim spName As String = DBO_SP_Return_String_NAME
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Dim Actual_Result As String = DB.SPGetSingleValue(Of String)(spName, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_Integer()
        Dim spName As String = DBO_SP_Return_Int_NAME
        Dim key As Integer = DBO_SP_Return_Int_VALUE
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Integer = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).MandatoryInteger
        Dim Actual_Result As Integer = DB.SPGetSingleValue(Of Integer)(spName, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_Boolean()
        Dim spName As String = DBO_SP_Return_Bool_NAME
        Dim key As Integer = DBO_SP_Return_Bool_VALUE
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Boolean = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Status
        Dim Actual_Result As Boolean = DB.SPGetBoolean(spName, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsFalse(Actual_Result)
    End Sub

    <TestMethod()>
    Public Sub SPGetSingleValue_NullDate()
        Dim spName As String = DBO_SP_Return_NullDate_NAME
        Dim key As Integer = DBO_SP_Return_NullDate_VALUE
        Dim parameter As New SqlParameter("@id", key)
        Dim Expected_Result As Date? = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).OptionalDate
        Dim Actual_Result As Date? = DB.SPGetSingleValue(Of Date?)(spName, parameter)
        Assert.AreEqual(Expected_Result, Actual_Result)
        Assert.IsNull(Actual_Result)
    End Sub

End Class