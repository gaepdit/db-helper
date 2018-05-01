Imports System.Data.SqlClient
Imports EpdIt.DBUtilities

' These functions call Stored Procedures that are written using a single OUTPUT parameter named "@return_value_argument".
' These are just a convenience to avoid having to code the output parameter and do the type conversion yourself.

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBooleanOutputParameter(spName As String,
                                                Optional parameter As SqlParameter = Nothing,
                                                Optional forceAddNullableParameters As Boolean = True
                                                ) As Boolean
        Return Convert.ToBoolean(SPGetOutputParameter(Of Boolean)(spName, parameter, forceAddNullableParameters))
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBooleanOutputParameter(spName As String,
                                                parameterArray As SqlParameter(),
                                                Optional forceAddNullableParameters As Boolean = True
                                                ) As Boolean
        Return Convert.ToBoolean(SPGetOutputParameter(Of Boolean)(spName, parameterArray, forceAddNullableParameters))
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetOutputParameter(Of T)(spName As String,
                                               Optional parameter As SqlParameter = Nothing,
                                               Optional forceAddNullableParameters As Boolean = True
                                               ) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetOutputParameter(Of T)(spName, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetOutputParameter(Of T)(spName As String,
                                               parameterArray As SqlParameter(),
                                               Optional forceAddNullableParameters As Boolean = True
                                               ) As T
        Dim result As Boolean = False
        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing Then
                    If forceAddNullableParameters Then
                        DBNullifyParameters(parameterArray)
                    End If
                    command.Parameters.AddRange(parameterArray)
                End If

                Dim outputParameter As New SqlParameter(SPReturnValueParameterName, TypeToSqlDbType(Of T)) With {
                    .Direction = ParameterDirection.Output,
                    .Size = Integer.MaxValue
                }
                command.Parameters.Add(outputParameter)

                command.Connection.Open()
                result = command.ExecuteNonQuery()
                command.Connection.Close()

                Return GetNullable(Of T)(command.Parameters(SPReturnValueParameterName).Value)
            End Using
        End Using
    End Function

    Private Function TypeToSqlDbType(Of T)() As SqlDbType
        ' See https://msdn.microsoft.com/en-us/library/yy6y35y8.aspx

        Select Case GetType(T)
            Case GetType(Boolean)
                Return SqlDbType.Bit
            Case GetType(Boolean?)
                Return SqlDbType.Bit
            Case GetType(Byte)
                Return SqlDbType.TinyInt
            Case GetType(Byte?)
                Return SqlDbType.TinyInt
            Case GetType(Byte())
                Return SqlDbType.VarBinary
            Case GetType(Char)
                Return SqlDbType.VarChar
            Case GetType(Date)
                Return SqlDbType.DateTime2
            Case GetType(Date?)
                Return SqlDbType.DateTime2
            Case GetType(DateTimeOffset)
                Return SqlDbType.DateTimeOffset
            Case GetType(Decimal)
                Return SqlDbType.Decimal
            Case GetType(Decimal?)
                Return SqlDbType.Decimal
            Case GetType(Double)
                Return SqlDbType.Float
            Case GetType(Double?)
                Return SqlDbType.Float
            Case GetType(Single)
                Return SqlDbType.Real
            Case GetType(Single?)
                Return SqlDbType.Real
            Case GetType(Guid)
                Return SqlDbType.UniqueIdentifier
            Case GetType(Short)
                Return SqlDbType.SmallInt
            Case GetType(Short?)
                Return SqlDbType.SmallInt
            Case GetType(Integer)
                Return SqlDbType.Int
            Case GetType(Integer?)
                Return SqlDbType.Int
            Case GetType(Long)
                Return SqlDbType.BigInt
            Case GetType(Long?)
                Return SqlDbType.BigInt
            Case GetType(Object)
                Return SqlDbType.Variant
            Case GetType(String)
                Return SqlDbType.NVarChar
            Case GetType(TimeSpan)
                Return SqlDbType.Time
            Case GetType(TimeSpan?)
                Return SqlDbType.Time
            Case Else
                Throw New ArgumentException("No SQL data type mapping defined for " & GetType(T).ToString & ". Use SPRunCommand instead.")
        End Select
    End Function

#Region " Deprecated methods "

    ''' <summary>
    ''' A synonym for SPGetBooleanOutputParameter. Kept for backward compatibility.
    ''' May be removed in a future version.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    <Obsolete("This method is deprecated, use SPGetBooleanOutputParameter instead.")>
    Public Function SPGetBooleanReturnValue(spName As String,
                                            Optional parameter As SqlParameter = Nothing,
                                            Optional forceAddNullableParameters As Boolean = True
                                            ) As Boolean
        Return SPGetBooleanOutputParameter(spName, parameter, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' A synonym for SPGetBooleanOutputParameter. Kept for backward compatibility.
    ''' May be removed in a future version.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A boolean value.</returns>
    <Obsolete("This method is deprecated, use SPGetBooleanOutputParameter instead.")>
    Public Function SPGetBooleanReturnValue(spName As String,
                                            parameterArray As SqlParameter(),
                                            Optional forceAddNullableParameters As Boolean = True
                                            ) As Boolean
        Return SPGetBooleanOutputParameter(spName, parameterArray, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' A synonym for SPGetOutputValue. Kept for backward compatibility.
    ''' May be removed in a future version.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    <Obsolete("This method is deprecated, use SPGetOutputValue instead.")>
    Public Function SPGetSingleReturnValue(Of T)(spName As String,
                                                 Optional parameter As SqlParameter = Nothing,
                                                 Optional forceAddNullableParameters As Boolean = True
                                                 ) As T
        Return SPGetOutputParameter(Of T)(spName, parameter, forceAddNullableParameters)
    End Function

    ''' <summary>
    ''' A synonym for SPGetOutputValue. Kept for backward compatibility.
    ''' May be removed in a future version.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <param name="forceAddNullableParameters">True to force sending DBNull.Value for parameters that evaluate to Nothing; false to allow default behavior of dropping such parameters.</param>
    ''' <returns>A value of the specified type.</returns>
    <Obsolete("This method is deprecated, use SPGetOutputValue instead.")>
    Public Function SPGetSingleReturnValue(Of T)(spName As String,
                                                 parameterArray As SqlParameter(),
                                                 Optional forceAddNullableParameters As Boolean = True
                                                 ) As T
        Return SPGetOutputParameter(Of T)(spName, parameterArray, forceAddNullableParameters)
    End Function

#End Region

End Class
