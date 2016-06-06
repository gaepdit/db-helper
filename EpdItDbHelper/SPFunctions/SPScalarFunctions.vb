Imports System.Data.SqlClient

' These functions call Stored Procedures and return the first field of the first record of the result

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String, Optional parameter As SqlParameter = Nothing) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameter))
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String, parameterArray As SqlParameter()) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameterArray))
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String, Optional parameter As SqlParameter = Nothing) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetSingleValue(Of T)(spName, parameterArray)
    End Function


    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String, parameterArray As SqlParameter()) As T
        Dim result As Object = Nothing

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
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

End Class
