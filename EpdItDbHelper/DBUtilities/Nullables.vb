Partial Public Class DBUtilities

    ''' <summary>
    ''' Generic function for converting database values to useable .NET values, handling DBNull values appropriately.
    ''' </summary>
    ''' <typeparam name="T">The expected data type.</typeparam>
    ''' <param name="obj">The database value to convert.</param>
    ''' <returns>If database value is DBNull, returns the default value for the requested data type; otherwise, returns the value unchanged.</returns>
    <DebuggerStepThrough()>
    Public Shared Function GetNullable(Of T)(ByVal obj As Object) As T
        ' http://stackoverflow.com/a/870771/212978
        ' http://stackoverflow.com/a/9953399/212978
        If obj Is Nothing OrElse IsDBNull(obj) OrElse obj.ToString = "null" Then
            ' returns the default value for the type
            Return CType(Nothing, T)
        Else
            Return CType(obj, T)
        End If
    End Function

    ''' <summary>
    ''' Converts a database value to a nullable DateTime object, handling DBNull values appropriately.
    ''' </summary>
    ''' <param name="obj">The database value to convert.</param>
    ''' <returns>If database value is DBNull or value cannot be converted to a DateTime, returns Nothing; otherwise, returns the value converted to a DateTime.</returns>
    <DebuggerStepThrough()>
    Public Shared Function GetNullableDateTimeFromString(ByVal obj As Object) As DateTime?
        Try

            If obj Is Nothing OrElse IsDBNull(obj) OrElse String.IsNullOrEmpty(obj) Then
                Return Nothing
            Else
                Dim newDate As New DateTime
                If DateTime.TryParse(obj, newDate) Then
                    Return newDate
                Else
                    Return Nothing
                End If
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
