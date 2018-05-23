Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

<TestClass()> Public Class TableFunctionsTests

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
    Public Sub GetDataTable_NoParams()
        Dim query As String = "SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]"
        Dim dt As DataTable = DB.GetDataTable(query)

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
    Public Sub GetDataTable_OneParam()
        Dim query As String = "SELECT * FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]"
        Dim parameter As New SqlParameter("@mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dt As DataTable = DB.GetDataTable(query, parameter)

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
    Public Sub GetDataRow_NoParams()
        Dim key As Integer = 2
        Dim query As String = "SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = " & key
        Dim dr As DataRow = DB.GetDataRow(query)

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
    Public Sub GetDataRow_OneParam()
        Dim key As Integer = 2
        Dim query As String = "SELECT ID, Name, Status, MandatoryInteger, MandatoryDate, OptionalInteger, OptionalDate FROM dbo.Things WHERE ID = @id"
        Dim parameter As New SqlParameter("@id", key)
        Dim dr As DataRow = DB.GetDataRow(query, parameter)

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
    Public Sub GetLookupDictionary_NoParams()
        Dim query As String = "SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = " & ThingData.ThingSelectionKey & " ORDER BY [ID]"
        Dim dict As Dictionary(Of Integer, String) = DB.GetLookupDictionary(query)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

    <TestMethod()>
    Public Sub GetLookupDictionary_OneParam()
        Dim query As String = "SELECT ID, Name FROM dbo.Things WHERE MandatoryInteger = @mandatoryInteger ORDER BY [ID]"
        Dim parameter As New SqlParameter("@mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dict As Dictionary(Of Integer, String) = DB.GetLookupDictionary(query, parameter)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

End Class