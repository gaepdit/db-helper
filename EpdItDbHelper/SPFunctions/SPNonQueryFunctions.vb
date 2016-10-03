Imports System.Data.SqlClient

' These functions call Stored Procedures using IN and/or OUT parameters.
' If successful, any OUT parameters are available to the calling procedure as 
' returned by the database.

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="rowsAffected">For UPDATE, INSERT, and DELETE statements, stores the number of rows affected by the command.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>True if the Stored Procedure ran successfully. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 Optional ByRef parameter As SqlParameter = Nothing,
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional forceAddNullableParameters As Boolean = False
                                 ) As Boolean
        rowsAffected = 0
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Dim result As Boolean = SPRunCommand(spName, parameterArray, rowsAffected, forceAddNullableParameters)
        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If
        Return result
    End Function

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameterArray">An SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>True if the Stored Procedure ran successfully. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional forceAddNullableParameters As Boolean = False
                                 ) As Boolean
        Dim success As Boolean = False

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If
                Dim returnParameter As New SqlParameter("@ReturnValue", SqlDbType.Int) With {
                    .Direction = ParameterDirection.ReturnValue
                }
                command.Parameters.Add(returnParameter)
                command.Connection.Open()
                rowsAffected = command.ExecuteNonQuery()
                command.Connection.Close()
                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    Dim newArray(command.Parameters.Count) As SqlParameter
                    command.Parameters.CopyTo(newArray, 0)
                    Array.Copy(newArray, parameterArray, parameterArray.Length)
                End If
                command.Parameters.Clear()
                success = (returnParameter.Value = 0)
            End Using
        End Using

        Return success
    End Function

End Class
