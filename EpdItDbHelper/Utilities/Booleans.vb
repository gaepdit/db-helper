Partial Public Class DBUtilities

    ''' <summary>
    ''' An enumeration of the different ways pseud-boolean values are stored in the database.
    ''' </summary>
    Public Enum BooleanDBConversionType
        TrueOrDBNull
        OneOrZero
    End Enum

    ''' <summary>
    ''' Converts a database value to a boolean according to the specified DB conversion type.
    ''' </summary>
    ''' <param name="value">The database value to convert.</param>
    ''' <param name="conversionType">A BooleanDBConversionType indicating how the value is stored in the database.</param>
    ''' <returns>A string that can be stored in the database.</returns>
    Public Shared Function ConvertDBValueToBoolean(value As String, conversionType As BooleanDBConversionType) As Boolean
        Select Case conversionType
            Case BooleanDBConversionType.TrueOrDBNull
                If value Is Nothing OrElse IsDBNull(value) OrElse value.ToString = "null" Then
                    Return False
                Else
                    Return Boolean.Parse(value) ' Will throw an exception if value is not equal to Boolean.TrueString
                End If

            Case BooleanDBConversionType.OneOrZero
                Return Convert.ToBoolean(Integer.Parse(value))

        End Select

        ' Fallback
        Return Boolean.Parse(value)
    End Function

    ''' <summary>
    ''' Converts a boolean value to a string that can be stored in the database according to the specified DB conversion type.
    ''' </summary>
    ''' <param name="value">The boolean value to convert.</param>
    ''' <param name="conversionType">A BooleanDBConversionType indicating how the value should be stored in the database.</param>
    ''' <returns>A string that can be stored in the database.</returns>
    Public Shared Function ConvertBooleanToDBValue(value As Boolean, conversionType As BooleanDBConversionType) As String
        Select Case conversionType
            Case BooleanDBConversionType.TrueOrDBNull
                If value Then
                    Return Boolean.TrueString
                Else
                    Return Nothing
                End If

            Case BooleanDBConversionType.OneOrZero
                If value Then
                    Return "1"
                Else
                    Return "0"
                End If

        End Select

        ' Fallback
        Return value.ToString
    End Function

    ''' <summary>
    ''' Utility for storing zeroed integers as null in the database
    ''' </summary>
    ''' <param name="i">The integer to store</param>
    ''' <returns>Nothing if i is equal to 0; otherwise, returns i.</returns>
    Public Shared Function StoreNothingIfZero(i As Integer) As Integer?
        If i = 0 Then
            Return Nothing
        Else
            Return i
        End If
    End Function

    ''' <summary>
    ''' Utility for storing zeroed decimals as null in the database
    ''' </summary>
    ''' <param name="i">The decimal to store</param>
    ''' <returns>Nothing if i is equal to 0; otherwise, returns i.</returns>
    Public Shared Function StoreNothingIfZero(i As Decimal) As Decimal?
        If i = 0 Then
            Return Nothing
        Else
            Return i
        End If
    End Function

End Class
