Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

<TestClass()>
Public Class SPTableFunctionsTests

    <TestInitialize>
    Public Sub InitializeTestDataTable()
        ThingData.SelectedThingsDataTable = New DataTable()
        ThingData.SelectedThingsDataTable.Columns.Add("ID")
        ThingData.SelectedThingsDataTable.Columns.Add("Name")
        ThingData.SelectedThingsDataTable.Columns.Add("Status")
        ThingData.SelectedThingsDataTable.Columns.Add("MandatoryInteger")
        ThingData.SelectedThingsDataTable.Columns.Add("MandatoryDate")

        For Each thing As Thing In ThingData.ThingsList
            If thing.MandatoryInteger = ThingData.ThingSelectionKey Then
                ThingData.SelectedThingsDataTable.Rows.Add({thing.ID, thing.Name, thing.Status, thing.MandatoryInteger, thing.MandatoryDate})
            End If
        Next
    End Sub

    <TestMethod()>
    Public Sub SPGetDataTable_NoParams()
        Dim spName As String = DBO_SP_Table_NoParameters_NAME
        Dim dt As DataTable = DB.SPGetDataTable(spName)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dt.Rows.Count)
        Assert.AreEqual(ThingData.ThingsList(0).ID, dt.Rows(0).Item("ID"))
        Assert.AreEqual(ThingData.ThingsList(0).Name, dt.Rows(0).Item("Name"))
        Assert.AreEqual(ThingData.ThingsList(0).Status, dt.Rows(0).Item("Status"))
        Assert.AreEqual(ThingData.ThingsList(0).MandatoryInteger, dt.Rows(0).Item("MandatoryInteger"))
        Assert.AreEqual(ThingData.ThingsList(0).MandatoryDate, dt.Rows(0).Item("MandatoryDate"))

        Assert.IsNull(GetNullable(Of Integer?)(dt.Rows(1).Item("OptionalInteger")))
        Assert.IsNull(GetNullable(Of Date?)(dt.Rows(1).Item("OptionalDate")))
    End Sub

    <TestMethod()>
    Public Sub SPGetDataTable_OneParam()
        Dim spName As String = DBO_SP_Table_OneParameter_NAME
        Dim parameter As New SqlParameter("mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dt As DataTable = DB.SPGetDataTable(spName, parameter)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dt.Rows.Count)
        Assert.AreEqual(ThingData.ThingsList(0).ID, dt.Rows(0).Item("ID"))
        Assert.AreEqual(ThingData.ThingsList(0).Name, dt.Rows(0).Item("Name"))
        Assert.AreEqual(ThingData.ThingsList(0).Status, dt.Rows(0).Item("Status"))
        Assert.AreEqual(ThingData.ThingsList(0).MandatoryInteger, dt.Rows(0).Item("MandatoryInteger"))
        Assert.AreEqual(ThingData.ThingsList(0).MandatoryDate, dt.Rows(0).Item("MandatoryDate"))

        Assert.IsNull(GetNullable(Of Integer?)(dt.Rows(1).Item("OptionalInteger")))
        Assert.IsNull(GetNullable(Of Date?)(dt.Rows(1).Item("OptionalDate")))
    End Sub

    <TestMethod>
    Public Sub SPGetDataRow_NoParam()
        Dim key As Integer = DBO_SP_Row_NoParameters_VALUE_int
        Dim spName As String = DBO_SP_Row_NoParameters_NAME
        Dim dr As DataRow = DB.GetDataRow(spName)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item("ID"))
        Assert.AreEqual(th.Name, dr.Item("Name"))
        Assert.AreEqual(th.Status, dr.Item("Status"))
        Assert.AreEqual(th.MandatoryInteger, dr.Item("MandatoryInteger"))
        Assert.AreEqual(th.MandatoryDate, dr.Item("MandatoryDate"))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item("OptionalInteger")))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item("OptionalDate")))
    End Sub

    <TestMethod>
    Public Sub SPGetDataRow_OneParam()
        Dim key As Integer = DBO_SP_Row_OneParameter_VALUE_int
        Dim spName As String = DBO_SP_Row_OneParameter_NAME
        Dim parameter As New SqlParameter("id", key)
        Dim dr As DataRow = DB.SPGetDataRow(spName, parameter)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item("ID"))
        Assert.AreEqual(th.Name, dr.Item("Name"))
        Assert.AreEqual(th.Status, dr.Item("Status"))
        Assert.AreEqual(th.MandatoryInteger, dr.Item("MandatoryInteger"))
        Assert.AreEqual(th.MandatoryDate, dr.Item("MandatoryDate"))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item("OptionalInteger")))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item("OptionalDate")))
    End Sub

    <TestMethod()>
    Public Sub SPGetLookupDictionary_NoParams()
        Dim spName As String = DBO_SP_Dictionary_NoParameters_NAME
        Dim dict As Dictionary(Of Integer, String) = DB.SPGetLookupDictionary(spName)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

    <TestMethod()>
    Public Sub SPGetLookupDictionary_OneParam()
        Dim spName As String = DBO_SP_Dictionary_OneParameter_NAME
        Dim parameter As New SqlParameter("mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dict As Dictionary(Of Integer, String) = DB.SPGetLookupDictionary(spName, parameter)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

End Class