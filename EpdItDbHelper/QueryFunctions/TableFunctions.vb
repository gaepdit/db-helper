Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

Partial Public Class DBHelper

#Region " DataRow "

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(query As String,
                               Optional parameter As SqlParameter = Nothing,
                               Optional forceAddNullableParameters As Boolean = True
                               ) As DataRow
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return GetDataRow(query, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataRow of values.</returns>
    Public Function GetDataRow(query As String,
                               parameterArray As SqlParameter(),
                               Optional forceAddNullableParameters As Boolean = True
                               ) As DataRow
        Dim resultTable As DataTable = GetDataTable(query, parameterArray, forceAddNullableParameters)
        If resultTable IsNot Nothing And resultTable.Rows.Count = 1 Then
            Return resultTable.Rows(0)
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region " DataTable "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(query As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As DataTable
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return GetDataTable(query, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable of values.</returns>
    Public Function GetDataTable(query As String,
                                 parameterArray As SqlParameter(),
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As DataTable
        Dim table As New DataTable

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(query, connection)
                command.CommandType = CommandType.Text
                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If
                Using adapter As New SqlDataAdapter(command)
                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                    command.Connection.Open()
                    adapter.Fill(table)
                    command.Connection.Close()
                End Using
                command.Parameters.Clear()
            End Using
        End Using

        Return table
    End Function

#End Region

#Region " Lookup Dictionary "

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function GetLookupDictionary(query As String,
                                        Optional parameter As SqlParameter = Nothing,
                                        Optional forceAddNullableParameters As Boolean = True
                                        ) As Dictionary(Of Integer, String)
        Dim d As New Dictionary(Of Integer, String)

        Dim dataTable As DataTable = GetDataTable(query, parameter, forceAddNullableParameters)

        For Each row As DataRow In dataTable.Rows
            d.Add(row.Item(0), row.Item(1))
        Next

        Return d
    End Function

#End Region

End Class
