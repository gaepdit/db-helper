Imports System.Data.SqlClient

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function GetBoolean(query As String,
                               Optional parameter As SqlParameter = Nothing
                               ) As Boolean
        Return Convert.ToBoolean(GetSingleValue(Of Boolean)(query, parameter))
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function GetBoolean(query As String,
                               parameterArray As SqlParameter()
                               ) As Boolean
        Return Convert.ToBoolean(GetSingleValue(Of Boolean)(query, parameterArray))
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function GetSingleValue(Of T)(query As String,
                                         Optional parameter As SqlParameter = Nothing
                                         ) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return GetSingleValue(Of T)(query, parameterArray)
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function GetSingleValue(Of T)(query As String,
                                         parameterArray As SqlParameter()
                                         ) As T
        Dim result As Object = Nothing

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(query, connection)
                command.CommandType = CommandType.Text
                If parameterArray IsNot Nothing Then
                    command.Parameters.AddRange(parameterArray)
                End If
                command.Connection.Open()
                result = command.ExecuteScalar()
                command.Connection.Close()
                command.Parameters.Clear()
            End Using
        End Using

        Return DBUtilities.GetNullable(Of T)(result)
    End Function

    ''' <summary>
    ''' Determines whether a value as indicated by the SQL query exists in the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter array to send.</param>
    ''' <returns>A boolean value signifying whether the indicated value exists.</returns>
    Public Function ValueExists(query As String,
                                Optional parameter As SqlParameter = Nothing
                                ) As Boolean
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return ValueExists(query, parameterArray)
    End Function

    ''' <summary>
    ''' Determines whether a value as indicated by the SQL query exists in the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A boolean value signifying whether the indicated value exists.</returns>
    Public Function ValueExists(query As String,
                                parameterArray As SqlParameter()
                                ) As Boolean
        Dim result As Object = Nothing

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(query, connection)
                command.CommandType = CommandType.Text
                If parameterArray IsNot Nothing Then
                    command.Parameters.AddRange(parameterArray)
                End If
                command.Connection.Open()
                result = command.ExecuteScalar()
                command.Connection.Close()
                command.Parameters.Clear()
            End Using
        End Using

        Return Not (result Is Nothing OrElse IsDBNull(result) OrElse result.ToString = "null")
    End Function

End Class
