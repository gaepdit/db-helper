Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

' These functions call Stored Procedures and return the first field of the first record of the result

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a scalar boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameter, forceAddNullableParameters))
    End Function

    ''' <summary>
    ''' Retrieves a scalar boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional forceAddNullableParameters As Boolean = True
                               ) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameterArray, forceAddNullableParameters))
    End Function

    ''' <summary>
    ''' Retrieves a scalar integer value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>An integer value.</returns>
    Public Function SPGetInteger(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As Integer
        Return SPGetSingleValue(Of Integer)(spName, parameter, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a scalar integer value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>An integer value.</returns>
    Public Function SPGetInteger(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional forceAddNullableParameters As Boolean = True
                               ) As Integer
        Return SPGetSingleValue(Of Integer)(spName, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a scalar string value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A string value.</returns>
    Public Function SPGetString(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional forceAddNullableParameters As Boolean = True
                                 ) As String
        Return SPGetSingleValue(Of String)(spName, parameter, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a scalar string value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A string value.</returns>
    Public Function SPGetString(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional forceAddNullableParameters As Boolean = True
                               ) As String
        Return SPGetSingleValue(Of String)(spName, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String,
                                           Optional parameter As SqlParameter = Nothing,
                                           Optional forceAddNullableParameters As Boolean = True
                                           ) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetSingleValue(Of T)(spName, parameterArray, forceAddNullableParameters)
    End Function


    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String,
                                           parameterArray As SqlParameter(),
                                           Optional forceAddNullableParameters As Boolean = True
                                           ) As T
        Dim result As Object = Nothing

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure
                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
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
