Imports System.Data.SqlClient

Partial Public Class DBHelper

    ''' <summary>
    ''' Loops through all SqlParameters in an array and if the Value of any evaluates to Nothing, sets the Value 
    ''' to DBNull.Value. This will force the parameter to be sent with the call to the database. 
    ''' 
    ''' (By default, parameters that evaluate to Nothing are removed from the database call, causing an 
    ''' error if the parameter is expected, even if null is an allowed value.)
    ''' </summary>
    ''' <param name="sqlParameter">An array of SqlParameter</param>
    Public Shared Sub DBNullifyParameters(ByRef sqlParameter() As SqlParameter)
        For Each param As SqlParameter In sqlParameter
            param.Value = If(param.Value, DBNull.Value)
        Next
    End Sub

End Class
