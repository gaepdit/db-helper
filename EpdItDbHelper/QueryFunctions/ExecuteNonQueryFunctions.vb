Imports System.Data.SqlClient

Partial Public Class DB

#Region " Run commands (ExecuteNonQuery) "

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <param name="failSilently">If true, suppresses error messages displayed to the user.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(ByVal query As String,
                                   Optional ByVal parameter As SqlParameter = Nothing,
                                   Optional ByRef rowsAffected As Integer = 0,
                                   Optional ByVal failSilently As Boolean = False
                                   ) As Boolean
        rowsAffected = 0
        Dim parameterArray As SqlParameter() = {parameter}
        Return RunCommand(query, parameterArray, rowsAffected, failSilently)
    End Function

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameters">An SqlParameter array to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <param name="failSilently">If true, suppresses error messages displayed to the user.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(ByVal query As String,
                                   ByVal parameters As SqlParameter(),
                                   Optional ByRef rowsAffected As Integer = 0,
                                   Optional ByVal failSilently As Boolean = False
                                   ) As Boolean
        rowsAffected = 0
        Dim queryList As New List(Of String)
        queryList.Add(query)

        Dim parametersList As New List(Of SqlParameter())
        parametersList.Add(parameters)

        Dim countList As New List(Of Integer)

        Dim result As Boolean = RunCommand(queryList, parametersList, countList, failSilently)

        If result AndAlso countList.Count > 0 Then rowsAffected = countList(0)

        Return result
    End Function

    ''' <summary>
    ''' Executes a set of SQL statements on the database.
    ''' </summary>
    ''' <param name="queryList">The SQL statements to execute.</param>
    ''' <param name="parametersList">A List of SqlParameter arrays to send.</param>
    ''' <param name="countList">A List of rows affected by each SQL statement.</param>
    ''' <param name="failSilently"></param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(ByVal queryList As List(Of String),
                                   ByVal parametersList As List(Of SqlParameter()),
                                   Optional ByRef countList As List(Of Integer) = Nothing,
                                   Optional ByVal failSilently As Boolean = False
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
                            command.Parameters.Clear()
                            command.CommandText = queryList(index)
                            command.Parameters.AddRange(parametersList(index))
                            Dim rowsAffected As Integer = command.ExecuteNonQuery()
                            countList.Insert(index, rowsAffected)
                        Next
                        transaction.Commit()
                    Catch ee As SqlException
                        success = False
                        countList.Clear()
                        Try
                            transaction.Rollback()
                        Catch
                        End Try
                    End Try

                    command.Connection.Close()
                Catch ee As SqlException
                    success = False
                Finally
                    If transaction IsNot Nothing Then transaction.Dispose()
                End Try

            End Using
        End Using

        Return success
    End Function

    Public Function RunCommandIgnoreErrors(query As String, parameters As SqlParameter()) As Boolean
        Try
            Using connection As New SqlConnection(ConnectionString)
                Using command As New SqlCommand(query, connection)
                    command.CommandType = CommandType.Text
                    command.Parameters.AddRange(parameters)
                    command.Connection.Open()
                    command.ExecuteScalar()
                    command.Connection.Close()
                End Using
            End Using
            Return True
        Catch
            Return False
        End Try
    End Function

#End Region

End Class
