Imports System.Data.SqlClient

Partial Public Class DBHelper

    Public Shared Sub DBNullifyParameters(ByRef sqlParameter() As SqlParameter)
        For Each param As SqlParameter In sqlParameter
            param.Value = If(param.Value, DBNull.Value)
        Next
    End Sub

End Class
