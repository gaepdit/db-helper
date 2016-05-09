Imports System.IO

Partial Public Class DBUtilities

    ''' <summary>
    ''' Reads a binary file into a Byte array.
    ''' </summary>
    ''' <param name="pathToFile">The path to the file to read.</param>
    ''' <returns>A Byte array.</returns>
    Public Shared Function ReadByteArrayFromFile(ByVal pathToFile As String) As Byte()
        Return File.ReadAllBytes(pathToFile)
    End Function

End Class
