Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

Partial Public Class DBHelper

#Region " DataRow "

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameter">An optional Oracle Parameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataRow.</returns>
    Public Function SPGetDataRow(spName As String,
                                 ByRef Optional parameter As SqlParameter = Nothing,
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As DataRow
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetDataRow(spName, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a single row of values from the database.
    ''' </summary>
    ''' <param name="spName">The Oracle Stored Procedure to call (SP must be a function that returns a REFCURSOR)</param>
    ''' <param name="parameterArray">An optional Oracle Parameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataRow</returns>
    Public Function SPGetDataRow(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As DataRow
        Dim resultTable As DataTable = SPGetDataTable(spName, parameterArray, forceAddNullableParameters)
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
    ''' <param name="spName">The Stored Procedure to call</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataTable(spName As String,
                                   Optional ByRef parameter As SqlParameter = Nothing,
                                   Optional forceAddNullableParameters As Boolean = True
                                   ) As DataTable
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Dim table As DataTable = SPGetDataTable(spName, parameterArray, forceAddNullableParameters)
        If table IsNot Nothing AndAlso parameterArray IsNot Nothing Then
            parameter = parameterArray(0)
        End If
        Return table
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The Stored Procedure to call</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataTable(spName As String,
                                   ByRef parameterArray As SqlParameter(),
                                   Optional forceAddNullableParameters As Boolean = True
                                   ) As DataTable
        If String.IsNullOrEmpty(spName) Then
            Return Nothing
        End If

        Dim table As New DataTable

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If
                Using adapter As New SqlDataAdapter(command)
                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                    adapter.Fill(table)
                    If parameterArray IsNot Nothing Then
                        command.Parameters.CopyTo(parameterArray, 0)
                    End If
                End Using
                command.Parameters.Clear()
            End Using
        End Using

        Return table
    End Function

#End Region

#Region " DataSet "

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The Stored Procedure to call</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataSet(spName As String,
                                   Optional ByRef parameter As SqlParameter = Nothing,
                                   Optional forceAddNullableParameters As Boolean = True
                                   ) As DataSet
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Dim ds As DataSet = SPGetDataSet(spName, parameterArray, forceAddNullableParameters)
        If ds IsNot Nothing AndAlso parameterArray IsNot Nothing Then
            parameter = parameterArray(0)
        End If
        Return ds
    End Function

    ''' <summary>
    ''' Retrieves a DataTable of values from the database.
    ''' </summary>
    ''' <param name="spName">The Stored Procedure to call</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A DataTable</returns>
    Public Function SPGetDataSet(spName As String,
                                   ByRef parameterArray As SqlParameter(),
                                   Optional forceAddNullableParameters As Boolean = True
                                   ) As DataSet
        If String.IsNullOrEmpty(spName) Then
            Return Nothing
        End If

        Dim ds As New DataSet

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If
                Using adapter As New SqlDataAdapter(command)
                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                    adapter.Fill(ds)
                    If parameterArray IsNot Nothing Then
                        command.Parameters.CopyTo(parameterArray, 0)
                    End If
                End Using
                command.Parameters.Clear()
            End Using
        End Using

        Return ds
    End Function

#End Region

#Region " Lookup Dictionary "

    ''' <summary>
    ''' Retrieves a dictionary of (integer -> string) values from the database
    ''' </summary>
    ''' <param name="spName">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A lookup dictionary.</returns>
    Public Function SPGetLookupDictionary(spName As String,
                                          ByRef Optional parameter As SqlParameter = Nothing,
                                          Optional forceAddNullableParameters As Boolean = True
                                          ) As Dictionary(Of Integer, String)
        Dim d As New Dictionary(Of Integer, String)

        Dim dataTable As DataTable = SPGetDataTable(spName, parameter, forceAddNullableParameters)

        For Each row As DataRow In dataTable.Rows
            d.Add(row.Item(0), row.Item(1))
        Next

        Return d
    End Function

#End Region

End Class
