Imports Microsoft.Data.SqlClient

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(query As String,
                               Optional parameter As SqlParameter = Nothing,
                               Optional ByRef rowsAffected As Integer = 0
                               ) As Boolean

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Return RunCommand(query, parameterArray, rowsAffected)
    End Function

    ''' <summary>
    ''' Executes a SQL statement on the database.
    ''' </summary>
    ''' <param name="query">The SQL statement to execute.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(query As String,
                               parameterArray As SqlParameter(),
                               Optional ByRef rowsAffected As Integer = 0
                               ) As Boolean

        Dim queryList As New List(Of String) From {query}
        Dim parameterArrayList As New List(Of SqlParameter()) From {parameterArray}
        Dim countList As New List(Of Integer)

        Dim result As Boolean = RunCommand(queryList, parameterArrayList, countList)

        If result AndAlso countList.Count > 0 Then rowsAffected = countList(0)

        Return result
    End Function

    ''' <summary>
    ''' Executes a set of SQL statements on the database.
    ''' </summary>
    ''' <param name="queryList">The SQL statements to execute.</param>
    ''' <param name="parametersList">A List of SqlParameter arrays to send.</param>
    ''' <param name="rowsAffectedList">A List of rows affected by each SQL statement.</param>
    ''' <returns>True if command ran successfully. Otherwise, false.</returns>
    Public Function RunCommand(queryList As List(Of String),
                               parametersList As List(Of SqlParameter()),
                               Optional ByRef rowsAffectedList As List(Of Integer) = Nothing
                               ) As Boolean

        Return QExecuteNonQuery(queryList, parametersList, rowsAffectedList)
    End Function

End Class
