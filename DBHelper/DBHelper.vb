Imports System.Data.SqlClient

Public Class DBHelper

    Private Property ConnectionString As String

    Public Sub New(connectionString As String)
        Me.ConnectionString = connectionString
    End Sub

    Private Function ReturnValueParameter() As SqlParameter
        Return New SqlParameter("@DbHelperReturnValue", SqlDbType.Int) With {.Direction = ParameterDirection.ReturnValue}
    End Function

    Public Class TooManyRecordsException
        Inherits Exception

        Private Const errorMessage As String = "Query returned more than one record."

        Public Sub New()
            MyBase.New(errorMessage)
        End Sub

        Public Sub New(auxMessage As String)
            MyBase.New(String.Format("{0} - {1}", errorMessage, auxMessage))
        End Sub

        Public Sub New(auxMessage As String, inner As Exception)
            MyBase.New(String.Format("{0} - {1}", errorMessage, auxMessage), inner)
        End Sub
    End Class

End Class
