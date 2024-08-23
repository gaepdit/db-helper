Imports System.Data
Imports Microsoft.Data.SqlClient
Imports GaEpd.DBUtilities
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class SPOutputParameterTests

    <TestMethod()>
    Public Sub SPOutput_Parameter()
        Dim spName As String = "SpOutParam"
        Dim key As Integer = 2
        Dim parameterArray As SqlParameter() = {
            New SqlParameter("@id", key),
            New SqlParameter("@name", SqlDbType.VarChar) With {
                .Direction = ParameterDirection.Output,
                .Size = 100
            }
        }

        Dim Result_Success As Boolean = DB.SPRunCommand(spName, parameterArray)
        Assert.IsTrue(Result_Success)

        Dim Actual_Output_Result As String = GetNullableString(parameterArray(1).Value)
        Dim Expected_Output_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Assert.AreEqual(Expected_Output_Result, Actual_Output_Result)
    End Sub

    <TestMethod()>
    Public Sub SPSelect_Input_and_Output_Parameter()
        Dim spName As String = "SpSelectInAndOutParam"
        Dim key As Integer = 2
        Dim parameterArray As SqlParameter() = {
            New SqlParameter("@id", key),
            New SqlParameter("@name", SqlDbType.VarChar) With {
                .Direction = ParameterDirection.Output,
                .Size = 100
            }
        }

        Dim QueryResult As Integer = DB.SPGetInteger(spName, parameterArray)
        Assert.AreEqual(ThingData.ThingSelectionKey, QueryResult)

        Dim Actual_Output_Result As String = GetNullableString(parameterArray(1).Value)
        Dim Expected_Output_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Assert.AreEqual(Expected_Output_Result, Actual_Output_Result)
    End Sub

    <TestMethod()>
    Public Sub SPSelect_Output_Parameter()
        Dim spName As String = "SpSelectOutParam"
        Dim key As Integer = 2
        Dim param = New SqlParameter("@name", SqlDbType.VarChar) With {
            .Direction = ParameterDirection.Output,
            .Size = 100
        }

        Dim QueryResult As Integer = DB.SPGetInteger(spName, param)
        Assert.AreEqual(ThingData.ThingSelectionKey, QueryResult)

        Dim Actual_Output_Result As String = GetNullableString(param.Value)
        Dim Expected_Output_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Assert.AreEqual(Expected_Output_Result, Actual_Output_Result)
    End Sub

    <TestMethod>
    Public Sub SPGetDataRow_and_Output_Param()
        Dim spName As String = "RowSpOutputParam"
        Dim key As Integer = 2
        Dim param As New SqlParameter("@name", SqlDbType.VarChar) With {
            .Direction = ParameterDirection.Output,
            .Size = 100
        }
        Dim dr As DataRow = DB.SPGetDataRow(spName, param)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item(NameOf(Thing.ID)))
        Assert.AreEqual(th.Name, dr.Item(NameOf(Thing.Name)))
        Assert.AreEqual(th.Status, dr.Item(NameOf(Thing.Status)))
        Assert.AreEqual(th.MandatoryInteger, dr.Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(th.MandatoryDate, dr.Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(th.OptionalInteger, GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(th.OptionalDate, GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))

        Dim Actual_Output_Result As String = GetNullableString(param.Value)
        Dim Expected_Output_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Assert.AreEqual(Expected_Output_Result, Actual_Output_Result)
    End Sub

    <TestMethod>
    Public Sub SPGetDataTable_Output_Param()
        Dim spName As String = "TableSpOutputParam"
        Dim key As Integer = 2
        Dim param As New SqlParameter("@dateOut", SqlDbType.Date) With {
            .Direction = ParameterDirection.Output
        }
        Dim dr As DataRow = DB.SPGetDataTable(spName, param).Rows.Find(key)

        Dim th As Thing = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key)

        Assert.AreEqual(th.ID, dr.Item(NameOf(Thing.ID)))
        Assert.AreEqual(th.Name, dr.Item(NameOf(Thing.Name)))
        Assert.AreEqual(th.Status, dr.Item(NameOf(Thing.Status)))
        Assert.AreEqual(th.MandatoryInteger, dr.Item(NameOf(Thing.MandatoryInteger)))
        Assert.AreEqual(th.MandatoryDate, dr.Item(NameOf(Thing.MandatoryDate)))
        Assert.AreEqual(th.OptionalInteger, GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.AreEqual(th.OptionalDate, GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))

        Assert.IsNull(GetNullable(Of Integer?)(dr.Item(NameOf(Thing.OptionalInteger))))
        Assert.IsNull(GetNullable(Of Date?)(dr.Item(NameOf(Thing.OptionalDate))))

        Dim Actual_Output_Result As Date = CDate(param.Value)
        Dim Expected_Output_Result As Date = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).MandatoryDate
        Assert.AreEqual(Expected_Output_Result, Actual_Output_Result)
    End Sub

    <TestMethod()>
    Public Sub SPReturnValue_and_Output_Param()
        Dim rowsAffected As Integer = 0
        Dim spName As String = "SpGetBooleanAndReturnAndOutputParam"
        Dim param As New SqlParameter("@dateOut", SqlDbType.Date) With {
            .Direction = ParameterDirection.Output
        }
        Dim Return_Value As Integer = DB.SPReturnValue(spName, param, rowsAffected:=rowsAffected)

        Assert.AreEqual(-1, rowsAffected)
        Assert.AreEqual(2, Return_Value)
        Assert.AreEqual(New Date(2019, 1, 1), param.Value)
    End Sub

    <TestMethod()>
    Public Sub SPGetBoolean_and_ReturnValue_and_Output_Param()
        Dim spName As String = "SpGetBooleanAndReturnAndOutputParam"
        Dim param As New SqlParameter("@dateOut", SqlDbType.Date) With {
            .Direction = ParameterDirection.Output
        }
        Dim returnValue As Integer = 0
        Dim result As Boolean = DB.SPGetBoolean(spName, param, returnValue:=returnValue)

        Assert.IsTrue(result)
        Assert.AreEqual(2, returnValue)
        Assert.AreEqual(New Date(2019, 1, 1), param.Value)
    End Sub

    <TestMethod()>
    Public Sub SPGetLookupDictionary_OutputParam()
        Dim spName As String = "DictSpOutputParam"
        Dim param As New SqlParameter("@dateOut", SqlDbType.Date) With {
            .Direction = ParameterDirection.Output
        }
        Dim dict As Dictionary(Of Integer, String) = DB.SPGetLookupDictionary(spName, param)

        Assert.AreEqual(ThingData.ThingsList.Where(Function(x) x.MandatoryInteger = ThingData.ThingSelectionKey).Count, dict.Count)
        Assert.AreEqual(ThingData.ThingsList(0).Name, dict(ThingData.ThingsList(0).ID))
        Assert.AreEqual(New Date(2019, 1, 1), param.Value)
    End Sub

End Class