Imports System.Data.SqlClient

Partial Public Class DBHelper

#Region " ByteArray "

    Public Function GetByteArray(ByVal query As String, ByVal parameterArray As SqlParameter()) As Byte()
        Throw New NotImplementedException

        'Using connection As New SqlConnection(ConnectionString)
        '    Using command As New SqlCommand(query, connection)
        '        command.CommandType = CommandType.Text
        '        Try
        '           If parameterArray IsNot Nothing Then
        '                command.Parameters.AddRange(parameterArray)
        '           End If

        '            command.Connection.Open()
        '            Dim dr As SqlDataReader = command.ExecuteReader()
        '            dr.Read()

        '            Dim length As Integer = dr.GetBytes(0, 0, Nothing, 0, Integer.MaxValue)
        '            Dim byteArray(length) As Byte
        '            dr.GetBytes(0, 0, byteArray, 0, length)

        '            dr.Dispose()
        '            command.Connection.Close()

        '            Return byteArray
        '        Catch ee As SqlException
        '            Return Nothing
        '        Catch ex As Exception
        '            Return Nothing
        '        Finally
        '            command.Parameters.Clear()
        '        End Try

        '    End Using
        'End Using
    End Function

#End Region

End Class
