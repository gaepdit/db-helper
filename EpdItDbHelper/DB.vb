﻿Public Class DB

    Private Property ConnectionString As String

    Public Sub New(DbConnectionString As String)
        ConnectionString = DbConnectionString
    End Sub

End Class
