Imports System.Data
Imports System.Data.SqlClient
Imports GaEpd.DBUtilities
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class SPTableFunctionsTests

    <TestInitialize>
    Public Sub InitializeTestDataTable()
        ThingData.SelectedThingsDataTable = New DataTable()
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.ID), Type.GetType("System.Int32"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.Name), Type.GetType("System.String"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.Status), Type.GetType("System.Boolean"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.MandatoryInteger), Type.GetType("System.Int32"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.MandatoryDate), Type.GetType("System.DateTime"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.OptionalInteger), Type.GetType("System.Int32"))
        ThingData.SelectedThingsDataTable.Columns.Add(NameOf(Thing.OptionalDate), Type.GetType("System.DateTime"))

        ThingData.AlternateThingsDataTable = New DataTable()
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.ID), Type.GetType("System.Int32"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.Name), Type.GetType("System.String"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.Status), Type.GetType("System.Boolean"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.MandatoryInteger), Type.GetType("System.Int32"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.MandatoryDate), Type.GetType("System.DateTime"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.OptionalInteger), Type.GetType("System.Int32"))
        ThingData.AlternateThingsDataTable.Columns.Add(NameOf(Thing.OptionalDate), Type.GetType("System.DateTime"))

        For Each thing As Thing In ThingData.ThingsList
            If thing.MandatoryInteger = ThingData.ThingSelectionKey Then
                ThingData.SelectedThingsDataTable.Rows.Add(thing.ID, thing.Name, thing.Status, thing.MandatoryInteger, thing.MandatoryDate, thing.OptionalInteger, thing.OptionalDate)
            Else
                ThingData.AlternateThingsDataTable.Rows.Add(thing.ID, thing.Name, thing.Status, thing.MandatoryInteger, thing.MandatoryDate, thing.OptionalInteger, thing.OptionalDate)
            End If
        Next
    End Sub

    <TestMethod()>
    Public Sub SPGetDataTable_NoParams()
        Dim spName As String = "TableSpNoParam"
        Dim dt As DataTable = DB.SPGetDataTable(spName)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dt.Rows.Count)
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.ID)), dt.Rows(0).Item(NameOf(Thing.ID)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Name)), dt.Rows(0).Item(NameOf(Thing.Name)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Status)), dt.Rows(0).Item(NameOf(Thing.Status)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryInteger)), dt.Rows(0).Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryDate)), dt.Rows(0).Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger)), dt.Rows(0).Item(NameOf(Thing.OptionalInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate)), dt.Rows(0).Item(NameOf(Thing.OptionalDate)))
        Assert.AreEqual(GetNullable(Of Integer?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger))), GetNullable(Of Integer?)(dt.Rows(0).Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(GetNullable(Of DateTime?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate))), GetNullable(Of DateTime?)(dt.Rows(0).Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dt.Rows(1).Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dt.Rows(1).Item(NameOf(Thing.OptionalDate))))
    End Sub

    <TestMethod()>
    Public Sub SPGetDataTable_OneParam()
        Dim spName As String = "TableSpOneParam"
        Dim parameter As New SqlParameter("mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dt As DataTable = DB.SPGetDataTable(spName, parameter)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.ID)), dt.Rows(0).Item(NameOf(Thing.ID)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Name)), dt.Rows(0).Item(NameOf(Thing.Name)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Status)), dt.Rows(0).Item(NameOf(Thing.Status)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryInteger)), dt.Rows(0).Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryDate)), dt.Rows(0).Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger)), dt.Rows(0).Item(NameOf(Thing.OptionalInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate)), dt.Rows(0).Item(NameOf(Thing.OptionalDate)))
        Assert.AreEqual(GetNullable(Of Integer?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger))), GetNullable(Of Integer?)(dt.Rows(0).Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(GetNullable(Of DateTime?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate))), GetNullable(Of DateTime?)(dt.Rows(0).Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dt.Rows(1).Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dt.Rows(1).Item(NameOf(Thing.OptionalDate))))
    End Sub

    <TestMethod>
    Public Sub SPGetDataRow_NoParam()
        Dim key As Integer = 2
        Dim spName As String = "RowSpNoParam"
        Dim dr As DataRow = DB.SPGetDataRow(spName)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item(NameOf(Thing.ID)))
        Assert.AreEqual(th.Name, dr.Item(NameOf(Thing.Name)))
        Assert.AreEqual(th.Status, dr.Item(NameOf(Thing.Status)))
        Assert.AreEqual(th.MandatoryInteger, dr.Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(th.MandatoryDate, dr.Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(th.OptionalInteger, GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(th.OptionalDate, GetNullable(Of DateTime?)(dr.Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))
    End Sub

    <TestMethod>
    Public Sub SPGetDataRow_OneParam()
        Dim key As Integer = 2
        Dim spName As String = "RowSpOneParam"
        Dim parameter As New SqlParameter("id", key)
        Dim dr As DataRow = DB.SPGetDataRow(spName, parameter)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item(NameOf(Thing.ID)))
        Assert.AreEqual(th.Name, dr.Item(NameOf(Thing.Name)))
        Assert.AreEqual(th.Status, dr.Item(NameOf(Thing.Status)))
        Assert.AreEqual(th.MandatoryInteger, dr.Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(th.MandatoryDate, dr.Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(th.OptionalInteger, GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(th.OptionalDate, GetNullable(Of DateTime?)(dr.Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))
    End Sub

    <TestMethod()>
    Public Sub SPGetLookupDictionary_NoParams()
        Dim spName As String = "DictSpNoParam"
        Dim dict As Dictionary(Of Integer, String) = DB.SPGetLookupDictionary(spName)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

    <TestMethod()>
    Public Sub SPGetLookupDictionary_OneParam()
        Dim spName As String = "DictSpOneParam"
        Dim parameter As New SqlParameter("mandatoryInteger", ThingData.ThingSelectionKey)
        Dim dict As Dictionary(Of Integer, String) = DB.SPGetLookupDictionary(spName, parameter)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
    End Sub

    <TestMethod()>
    Public Sub SPGetDataSet_NoParams()
        Dim spName As String = "DataSetSpNoParam"
        Dim ds As DataSet = DB.SPGetDataSet(spName)

        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows.Count, ds.Tables(0).Rows.Count)
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.ID)), ds.Tables(0).Rows(0).Item(NameOf(Thing.ID)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Name)), ds.Tables(0).Rows(0).Item(NameOf(Thing.Name)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.Status)), ds.Tables(0).Rows(0).Item(NameOf(Thing.Status)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryInteger)), ds.Tables(0).Rows(0).Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryDate)), ds.Tables(0).Rows(0).Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger)), ds.Tables(0).Rows(0).Item(NameOf(Thing.OptionalInteger)))
        Assert.AreEqual(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate)), ds.Tables(0).Rows(0).Item(NameOf(Thing.OptionalDate)))
        Assert.AreEqual(GetNullable(Of Integer?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger))), GetNullable(Of Integer?)(ds.Tables(0).Rows(0).Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(GetNullable(Of DateTime?)(ThingData.SelectedThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate))), GetNullable(Of DateTime?)(ds.Tables(0).Rows(0).Item(NameOf(Thing.OptionalDate))))

        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.ID)), ds.Tables(1).Rows(0).Item(NameOf(Thing.ID)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.Name)), ds.Tables(1).Rows(0).Item(NameOf(Thing.Name)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.Status)), ds.Tables(1).Rows(0).Item(NameOf(Thing.Status)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryInteger)), ds.Tables(1).Rows(0).Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.MandatoryDate)), ds.Tables(1).Rows(0).Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger)), ds.Tables(1).Rows(0).Item(NameOf(Thing.OptionalInteger)))
        Assert.AreEqual(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate)), ds.Tables(1).Rows(0).Item(NameOf(Thing.OptionalDate)))
        Assert.AreEqual(GetNullable(Of Integer?)(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalInteger))), GetNullable(Of Integer?)(ds.Tables(1).Rows(0).Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(GetNullable(Of DateTime?)(ThingData.AlternateThingsDataTable.Rows(0).Item(NameOf(Thing.OptionalDate))), GetNullable(Of DateTime?)(ds.Tables(1).Rows(0).Item(NameOf(Thing.OptionalDate))))
    End Sub

End Class