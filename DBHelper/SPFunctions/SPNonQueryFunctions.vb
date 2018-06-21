Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

' These functions call Stored Procedures using IN and/or OUT parameters.
' If successful, any OUT parameters are available to the calling procedure as 
' returned by the database.

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameter">(Optional) SqlParameter to send.</param>
    ''' <param name="rowsAffected">(Optional) Output parameter that stores the number of rows affected.</param>
    ''' <param name="forceAddNullableParameters">(Optional) True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <param name="returnValue">(Optional) Output parameter that stores the return value.</param>
    ''' <returns>True if the Stored Procedure returns a value of 0. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 Optional ByRef parameter As SqlParameter = Nothing,
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional forceAddNullableParameters As Boolean = True,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        returnValue = SPReturnValue(spName, parameter, rowsAffected, forceAddNullableParameters)

        Return (returnValue = 0)
    End Function

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameterArray">SqlParameter array to send.</param>
    ''' <param name="rowsAffected">(Optional) Output parameter that stores the number of rows affected.</param>
    ''' <param name="forceAddNullableParameters">(Optional) True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <param name="returnValue">(Optional) Output parameter that stores the return value.</param>
    ''' <returns>True if the Stored Procedure returns a value of 0. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 ByRef parameterArray As SqlParameter(),
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional forceAddNullableParameters As Boolean = True,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        returnValue = SPReturnValue(spName, parameterArray, rowsAffected, forceAddNullableParameters)

        Return (returnValue = 0)
    End Function

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameter">(optional) SqlParameter to send.</param>
    ''' <param name="rowsAffected">(Optional) Output parameter that stores the number of rows affected.</param>
    ''' <param name="forceAddNullableParameters">(Optional) True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>Integer RETURN value from the Stored Procedure.</returns>
    Public Function SPReturnValue(spName As String,
                                  Optional ByRef parameter As SqlParameter = Nothing,
                                  Optional ByRef rowsAffected As Integer = 0,
                                  Optional forceAddNullableParameters As Boolean = True
                                  ) As Integer

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim result As Integer = SPReturnValue(spName, parameterArray, rowsAffected, forceAddNullableParameters)

        If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
            parameter = parameterArray(0)
        End If

        Return result
    End Function

    ''' <summary>
    ''' Executes a Stored Procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the Stored Procedure to execute.</param>
    ''' <param name="parameterArray">SqlParameter array to send.</param>
    ''' <param name="rowsAffected">(Optional) Output parameter that stores the number of rows affected.</param>
    ''' <param name="forceAddNullableParameters">(Optional) True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>Integer RETURN value from the Stored Procedure.</returns>
    Public Function SPReturnValue(spName As String,
                                  ByRef parameterArray As SqlParameter(),
                                  Optional ByRef rowsAffected As Integer = 0,
                                  Optional forceAddNullableParameters As Boolean = True
                                  ) As Integer

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)

                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If

                Dim returnParameter As New SqlParameter("@EpdItReturnValue", SqlDbType.Int) With {
                    .Direction = ParameterDirection.ReturnValue
                }

                command.Parameters.Add(returnParameter)
                command.Connection.Open()
                rowsAffected = command.ExecuteNonQuery()
                command.Connection.Close()
                command.Parameters.Remove(returnParameter)

                If parameterArray IsNot Nothing AndAlso parameterArray.Count > 0 Then
                    Dim newArray(command.Parameters.Count) As SqlParameter
                    command.Parameters.CopyTo(newArray, 0)
                    Array.Copy(newArray, parameterArray, parameterArray.Length)
                End If

                command.Parameters.Clear()

                Return returnParameter.Value
            End Using
        End Using
    End Function

End Class
