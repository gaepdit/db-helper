Imports System.Data.SqlClient

Partial Public Class DBUtilities

    ''' <summary>
    ''' Generic function for converting database values to useable .NET values, handling DBNull values appropriately.
    ''' </summary>
    ''' <typeparam name="T">The expected data type.</typeparam>
    ''' <param name="obj">The database value to convert.</param>
    ''' <returns>If database value is DBNull, returns the default value for the requested data type; otherwise, returns the value unchanged.</returns>
    <DebuggerStepThrough()>
    Public Shared Function GetNullable(Of T)(obj As Object) As T
        ' http://stackoverflow.com/a/870771/212978
        ' http://stackoverflow.com/a/9953399/212978
        If obj Is Nothing OrElse IsDBNull(obj) OrElse obj.ToString = "null" Then
            ' returns the default value for the type
            Return Nothing
        Else
            Return CType(obj, T)
        End If
    End Function

    ''' <summary>
    ''' Generic function for converting database values to useable .NET strings, handling DBNull values appropriately.
    ''' </summary>
    ''' <param name="obj">The database value to convert.</param>
    ''' <returns>If database value is DBNull, returns the default value for the requested data type; otherwise, returns the value unchanged.</returns>
    <DebuggerStepThrough()>
    Public Shared Function GetNullableString(obj As Object) As String
        If obj Is Nothing OrElse IsDBNull(obj) OrElse obj.ToString = "null" Then
            ' returns the default value for the type
            Return Nothing
        Else
            Return CStr(obj)
        End If
    End Function

    ''' <summary>
    ''' Converts a database value to a nullable DateTime object, handling DBNull values appropriately.
    ''' </summary>
    ''' <param name="obj">The database value to convert.</param>
    ''' <returns>If database value is DBNull or value cannot be converted to a DateTime, returns Nothing; otherwise, returns the value converted to a DateTime.</returns>
    <DebuggerStepThrough()>
    Public Shared Function GetNullableDateTime(obj As Object) As DateTime?
        Try
            If obj Is Nothing OrElse IsDBNull(obj) OrElse String.IsNullOrEmpty(obj.ToString) Then
                Return Nothing
            Else
                Dim newDate As New DateTime
                If Date.TryParse(obj.ToString, newDate) Then
                    Return newDate
                Else
                    Return Nothing
                End If
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Loops through all SqlParameters in an array and if the Value of any evaluates to Nothing, sets the Value 
    ''' to DBNull.Value. This will force the parameter to be sent with the call to the database. 
    ''' 
    ''' (By default, parameters that evaluate to Nothing are removed from the database call, causing an 
    ''' error if the parameter is expected, even if null is an allowed value.)
    ''' </summary>
    ''' <param name="sqlParameter">An array of SqlParameter</param>
    Friend Shared Sub DBNullifyParameters(ByRef sqlParameter() As SqlParameter)
        For Each param As SqlParameter In sqlParameter
            param.Value = If(param.Value, DBNull.Value)
        Next
    End Sub

End Class
