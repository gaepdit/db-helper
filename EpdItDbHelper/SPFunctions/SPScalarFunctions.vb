Imports System.Data.SqlClient

Partial Public Class DBHelper

#Region " Read (Scalar) "

    Public Function SPGetInteger(spName As String) As Integer
        Dim result As Integer
        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.Text
                Dim returnValue As SqlParameter = command.Parameters.Add("@RETURN_VALUE", SqlDbType.Int)
                returnValue.Direction = ParameterDirection.ReturnValue
                command.Connection.Open()
                command.ExecuteNonQuery()
                result = command.Parameters("@RETURN_VALUE").Value
                command.Connection.Close()
            End Using
        End Using

        Return DBUtilities.GetNullable(Of Integer)(result)
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="failSilently">If true, OracleExceptions will be suppressed.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(ByVal query As String, Optional ByVal parameter As SqlParameter = Nothing, Optional ByVal failSilently As Boolean = False) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(query, parameter, failSilently))
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="failSilently">If true, OracleExceptions will be suppressed.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(ByVal query As String, ByVal parameterArray As SqlParameter(), Optional ByVal failSilently As Boolean = False) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(query, parameterArray, failSilently))
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="failSilently">If true, OracleExceptions will be suppressed.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(ByVal query As String, Optional ByVal parameter As SqlParameter = Nothing, Optional ByVal failSilently As Boolean = False) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetSingleValue(Of T)(query, parameterArray, failSilently)
    End Function


    ''' <summary>
    ''' Retrieves a single value of the specified type from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="failSilently">If true, OracleExceptions will be suppressed.</param>
    ''' <returns>A value of the specified type.</returns>
    Private Function SPGetSingleValue(Of T)(ByVal query As String, ByVal parameterArray As SqlParameter(), Optional ByVal failSilently As Boolean = False) As T
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
            End Using
        End Using

        Return DBUtilities.GetNullable(Of T)(result)
    End Function

#End Region

End Class
