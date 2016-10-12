Imports System.Data.SqlClient
Imports Microsoft.SqlServer.Server

Partial Public Class DBUtilities

    ''' <summary>
    ''' Converts an IEnumerable of T to an IEnumerable of SqlDataRecord.
    ''' </summary>
    ''' <typeparam name="T">The IEnumerable type.</typeparam>
    ''' <param name="values">The IEnumerable to convert.</param>
    ''' <returns>An IEnumerable of SqlDataRecord.</returns>
    ''' <remarks>See http://stackoverflow.com/a/337864/212978 </remarks>
    Private Shared Function SqlDataRecords(Of T)(values As IEnumerable(Of T), dbColumnName As String) As IEnumerable(Of SqlDataRecord)
        If values Is Nothing OrElse Not values.Any() Then Return Nothing

        Dim metadata As SqlMetaData = SqlMetaData.InferFromValue(values.First(), dbColumnName)
        Return values.Select(Function(v)
                                 Dim r As New SqlDataRecord(metadata)
                                 r.SetValues(v)
                                 Return r
                             End Function)
    End Function

    ''' <summary>
    ''' Returns a Structured (table-valued) SQL parameter using the IEnumerable values provided
    ''' </summary>
    ''' <typeparam name="T">The IEnumerable type.</typeparam>
    ''' <param name="parameterName">The SqlParameter name.</param>
    ''' <param name="values">The IEnumerable to set as the value of the SqlParameter</param>
    ''' <param name="dbTableTypeName">The name of the Table Type in the database</param>
    ''' <returns>A table-valued SqlParameter of type T, containing the supplied values.</returns>
    ''' <remarks>See http://stackoverflow.com/a/337864/212978 </remarks>
    Public Shared Function TvpSqlParameter(Of T)(parameterName As String, values As IEnumerable(Of T), dbTableTypeName As String, dbColumnName As String) As SqlParameter
        Return New SqlParameter(parameterName, SqlDbType.Structured) With {
            .Value = SqlDataRecords(Of T)(values, dbColumnName),
            .TypeName = dbTableTypeName
        }
    End Function

End Class
