Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(query As String,
                               Optional parameter As SqlParameter = Nothing,
                               Optional ByRef rowsAffected As Integer = 0,
                               Optional forceAddNullableParameters As Boolean = True
                               ) As Boolean
        rowsAffected = 0
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return RunCommand(query, parameterArray, rowsAffected, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(query As String,
                               parameterArray As SqlParameter(),
                               Optional ByRef rowsAffected As Integer = 0,
                               Optional forceAddNullableParameters As Boolean = True
                               ) As Boolean
        rowsAffected = 0
        Dim queryList As New List(Of String)
        queryList.Add(query)

        Dim parameterArrayList As New List(Of SqlParameter()) From {parameterArray}

        Dim countList As New List(Of Integer)

        Dim result As Boolean = RunCommand(queryList, parameterArrayList, countList, forceAddNullableParameters)

        If result AndAlso countList.Count > 0 Then rowsAffected = countList(0)

        Return result
    End Function

    ''' <summary>
    ''' Executes a set of SQL statements on the database.
    ''' </summary>
    ''' <param name="queryList">The SQL statements to execute.</param>
    ''' <param name="parametersList">A List of SqlParameter arrays to send.</param>
    ''' <param name="countList">A List of rows affected by each SQL statement.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(queryList As List(Of String),
                               parametersList As List(Of SqlParameter()),
                               Optional ByRef countList As List(Of Integer) = Nothing,
                               Optional forceAddNullableParameters As Boolean = True
                               ) As Boolean
        If countList Is Nothing Then countList = New List(Of Integer)
        countList.Clear()
        If queryList.Count <> parametersList.Count Then Return False
        Dim success As Boolean = True

        Using connection As New SqlConnection(ConnectionString)
            Using command As SqlCommand = connection.CreateCommand
                command.CommandType = CommandType.Text
                Dim transaction As SqlTransaction = Nothing

                Try
                    command.Connection.Open()
                    transaction = connection.BeginTransaction
                    command.Transaction = transaction

                    Try
                        For index As Integer = 0 To queryList.Count - 1
                            command.CommandText = queryList(index)
                            If parametersList(index) IsNot Nothing Then
                                If forceAddNullableParameters Then
                                    DBNullifyParameters(parametersList(index))
                                End If
                                command.Parameters.AddRange(parametersList(index))
                            End If
                            Dim rowsAffected As Integer = command.ExecuteNonQuery()
                            countList.Insert(index, rowsAffected)
                            command.Parameters.Clear()
                        Next
                        transaction.Commit()
                    Catch ee As SqlException
                        success = False
                        countList.Clear()
                        transaction.Rollback()
                        Throw
                    End Try

                    command.Connection.Close()
                Finally
                    If transaction IsNot Nothing Then transaction.Dispose()
                End Try

            End Using
        End Using

        Return success
    End Function

End Class
