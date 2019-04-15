Imports System.Data.SqlClient
Imports EpdIt

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

        Dim Actual_Result As String = DBUtilities.GetNullableString(parameterArray(1).Value)
        Dim Expected_Result As String = ThingData.ThingsList.Find(Function(Thing) Thing.ID = key).Name
        Assert.AreEqual(Expected_Result, Actual_Result)
    End Sub

End Class