Imports System.Data.SqlClient
Imports System.IO

Partial Public Class DB

#Region " Read (ByteArray) "

    Public Function SaveBinaryFileFromDB(filePath As String, query As String, Optional ByVal parameter As SqlParameter = Nothing) As Boolean
        Dim parameterArray As SqlParameter() = {parameter}
        Return SaveBinaryFileFromDB(filePath, query, parameterArray)
    End Function

    Public Function SaveBinaryFileFromDB(filePath As String, query As String, ByVal parameterArray As SqlParameter()) As Boolean
        Dim byteArray As Byte() = GetByteArrayFromBlob(query, parameterArray)

        Try
            Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write)
                Using bw As New BinaryWriter(fs)
                    bw.Write(byteArray)
                End Using ' bw
            End Using ' fs

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function GetByteArrayFromBlob(ByVal query As String, ByVal parameterArray As SqlParameter()) As Byte()
        Dim success As Boolean = True

        Using connection As New SqlConnection(ConnectionString)
            Using command As New SqlCommand(query, connection)
                command.CommandType = CommandType.Text
                command.Parameters.AddRange(parameterArray)

                Try
                    command.Connection.Open()
                    Dim dr As SqlDataReader = command.ExecuteReader()
                    dr.Read()

                    Dim length As Integer = dr.GetBytes(0, 0, Nothing, 0, Integer.MaxValue)
                    Dim byteArray(length) As Byte
                    dr.GetBytes(0, 0, byteArray, 0, length)

                    dr.Dispose()
                    command.Connection.Close()

                    Return byteArray
                Catch ee As SqlException
                    success = False
                    Return Nothing
                Catch ex As Exception
                    success = False
                    Return Nothing
                End Try

            End Using
        End Using
    End Function

#End Region

End Class
