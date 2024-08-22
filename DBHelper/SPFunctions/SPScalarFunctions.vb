Imports Microsoft.Data.SqlClient

' These functions call Stored Procedures and return the first field of the first record of the result

Partial Public Class DBHelper

    ''' <summary>
    ''' Retrieves a scalar boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameter, returnValue))
    End Function

    ''' <summary>
    ''' Retrieves a scalar boolean value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A boolean value.</returns>
    Public Function SPGetBoolean(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        Return Convert.ToBoolean(SPGetSingleValue(Of Boolean)(spName, parameterArray, returnValue))
    End Function

    ''' <summary>
    ''' Retrieves a scalar integer value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>An integer value.</returns>
    Public Function SPGetInteger(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Integer

        Return SPGetSingleValue(Of Integer)(spName, parameter, returnValue)
    End Function

    ''' <summary>
    ''' Retrieves a scalar integer value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>An integer value.</returns>
    Public Function SPGetInteger(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Integer

        Return SPGetSingleValue(Of Integer)(spName, parameterArray, returnValue)
    End Function

    ''' <summary>
    ''' Retrieves a scalar string value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">An optional SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A string value.</returns>
    Public Function SPGetString(spName As String,
                                Optional parameter As SqlParameter = Nothing,
                                Optional ByRef returnValue As Integer = Nothing
                                ) As String

        Return SPGetSingleValue(Of String)(spName, parameter, returnValue)
    End Function

    ''' <summary>
    ''' Retrieves a scalar string value from the database by calling a stored procedure.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A string value.</returns>
    Public Function SPGetString(spName As String,
                                parameterArray As SqlParameter(),
                                Optional ByRef returnValue As Integer = Nothing
                                ) As String

        Return SPGetSingleValue(Of String)(spName, parameterArray, returnValue)
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <typeparam name="T">The expected type of the value retrieved from the database.</typeparam>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">A SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String,
                                           Optional parameter As SqlParameter = Nothing,
                                           Optional ByRef returnValue As Integer = Nothing
                                           ) As T

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim result As T = SPGetSingleValue(Of T)(spName, parameterArray, returnValue)

        Return result
    End Function

    ''' <summary>
    ''' Retrieves a single value of the specified type from the database by calling a stored procedure.
    ''' </summary>
    ''' <typeparam name="T">The expected type of the value retrieved from the database.</typeparam>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>A value of the specified type.</returns>
    Public Function SPGetSingleValue(Of T)(spName As String,
                                           parameterArray As SqlParameter(),
                                           Optional ByRef returnValue As Integer = Nothing
                                           ) As T

        Dim result As Object = SPExecuteScalar(spName, parameterArray, returnValue)

        Return DBUtilities.GetNullable(Of T)(result)
    End Function

End Class
