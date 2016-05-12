Imports System.Data.SqlClient

' These functions call Stored Procedures that are written using a single OUTPUT parameter named "@return_value_argument".
' These are just a convenience to avoid having to code the output parameter and do the type conversion yourself.

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(query As String, Optional parameter As SqlParameter = Nothing) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(query, parameter))
    End Function

    ''' <summary>
    ''' Retrieves a boolean value from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(query As String, parameterArray As SqlParameter()) As Boolean
        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(query, parameterArray))
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database.
    ''' </summary>
    ''' <param name="query">The SQL query to send.</param>
    ''' <param name="parameter">An optional SqlParameter to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(query As String, Optional parameter As SqlParameter = Nothing) As T
        Dim parameterArray As SqlParameter() = Nothing
        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If
        Return SPGetSingleValue(Of T)(query, parameterArray)
    End Function


    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The stored procedure to call.</param>
    ''' <param name="parameterArray">An optional SqlParameter array to send.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String, parameterArray As SqlParameter()) As T
        Dim result As Boolean = False
        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(spName, connection)
                command.CommandType = CommandType.StoredProcedure

                If parameterArray IsNot Nothing Then
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

                Return DBUtilities.GetNullable(Of T)(command.Parameters(SPReturnValueParameterName).Value)
            End Using
        End Using
    End Function

    Private Function TypeToSqlDbType(Of T)() As SqlDbType
        ' See https://msdn.microsoft.com/en-us/library/yy6y35y8.aspx

        Select Case GetType(T)
            Case GetType(Boolean)
                Return SqlDbType.Bit
            Case GetType(Byte)
                Return SqlDbType.TinyInt
            Case GetType(Byte())
                Return SqlDbType.VarBinary
            Case GetType(Char)
                Return SqlDbType.VarChar
            Case GetType(Date)
                Return SqlDbType.DateTime
            Case GetType(DateTimeOffset)
                Return SqlDbType.DateTimeOffset
            Case GetType(Decimal)
                Return SqlDbType.Decimal
            Case GetType(Double)
                Return SqlDbType.Float
            Case GetType(Single)
                Return SqlDbType.Real
            Case GetType(Guid)
                Return SqlDbType.UniqueIdentifier
            Case GetType(Short)
                Return SqlDbType.SmallInt
            Case GetType(Integer)
                Return SqlDbType.Int
            Case GetType(Long)
                Return SqlDbType.BigInt
            Case GetType(Object)
                Return SqlDbType.Variant
            Case GetType(String)
                Return SqlDbType.NVarChar
            Case GetType(TimeSpan)
                Return SqlDbType.Time
            Case Else
                Throw New ArgumentException("No SQL data type mapping defined for " & GetType(T).ToString)
        End Select
    End Function

End Class
