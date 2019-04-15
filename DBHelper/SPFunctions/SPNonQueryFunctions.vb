Imports System.Data.SqlClient

' These functions call Stored Procedures using IN and/or OUT parameters.
' If successful, any OUT parameters are available to the calling procedure as 
' returned by the database.

Partial Public Class DBHelper

    ''' <summary>
    ''' Executes a non-query stored procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">A SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="rowsAffected">Optional output parameter that stores the number of rows affected.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>True if the stored procedure returns a value of 0. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 Optional parameter As SqlParameter = Nothing,
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        returnValue = SPReturnValue(spName, parameter, rowsAffected)

        Return (returnValue = 0)
    End Function

    ''' <summary>
    ''' Executes a non-query stored procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="rowsAffected">Optional output parameter that stores the number of rows affected.</param>
    ''' <param name="returnValue">Optional output parameter that stores the RETURN value of the stored procedure.</param>
    ''' <returns>True if the stored procedure returns a value of 0. Otherwise, false.</returns>
    Public Function SPRunCommand(spName As String,
                                 parameterArray As SqlParameter(),
                                 Optional ByRef rowsAffected As Integer = 0,
                                 Optional ByRef returnValue As Integer = Nothing
                                 ) As Boolean

        returnValue = SPReturnValue(spName, parameterArray, rowsAffected)

        Return (returnValue = 0)
    End Function

    ''' <summary>
    ''' Executes a non-query stored procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameter">A SqlParameter value. The value may be modified by the stored produre if it is an output parameter.</param>
    ''' <param name="rowsAffected">Optional output parameter that stores the number of rows affected.</param>
    ''' <returns>Integer RETURN value from the Stored Procedure.</returns>
    Public Function SPReturnValue(spName As String,
                                  Optional parameter As SqlParameter = Nothing,
                                  Optional ByRef rowsAffected As Integer = 0
                                  ) As Integer

        Dim parameterArray As SqlParameter() = Nothing

        If parameter IsNot Nothing Then
            parameterArray = {parameter}
        End If

        Dim returnValue As Integer = SPReturnValue(spName, parameterArray, rowsAffected)

        Return returnValue
    End Function

    ''' <summary>
    ''' Executes a non-query stored procedure on the database.
    ''' </summary>
    ''' <param name="spName">The name of the stored procedure to execute.</param>
    ''' <param name="parameterArray">An array of SqlParameter values. The array may be modified by the stored produre if it includes output parameters.</param>
    ''' <param name="rowsAffected">Optional output parameter that stores the number of rows affected.</param>
    ''' <returns>Integer RETURN value from the Stored Procedure.</returns>
    Public Function SPReturnValue(spName As String,
                                  parameterArray As SqlParameter(),
                                  Optional ByRef rowsAffected As Integer = 0
                                  ) As Integer

        Dim returnValue As Integer

        rowsAffected = SPExecuteNonQuery(spName, parameterArray, returnValue)

        Return returnValue
    End Function

End Class
