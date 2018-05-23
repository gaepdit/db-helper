Imports EpdIt.DBUtilities

<TestClass()>
Public Class NullableUtilitiesTests

    <TestMethod>
    Public Sub GetNullable_ToString_FromDBNull()
        Dim result As String = GetNullable(Of String)(DBNull.Value)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableString_FromDBNull()
        Dim result As String = GetNullableString(DBNull.Value)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToBool_FromDBNull()
        Dim result As Boolean = GetNullable(Of Boolean)(DBNull.Value)
        Assert.AreEqual(False, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToInteger_FromDBNull()
        Dim result As Integer = GetNullable(Of Integer)(DBNull.Value)
        Assert.AreEqual(0, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToString_FromNull()
        Dim result As String = GetNullable(Of String)(Nothing)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableString_FromNull()
        Dim result As String = GetNullableString(Nothing)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToBool_FromNull()
        Dim result As Boolean = GetNullable(Of Boolean)(Nothing)
        Assert.AreEqual(False, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToInteger_FromNull()
        Dim result As Integer = GetNullable(Of Integer)(Nothing)
        Assert.AreEqual(0, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToString_FromString()
        Dim result As String = GetNullable(Of String)("test")
        Assert.AreEqual("test", result)
    End Sub

    <TestMethod>
    Public Sub GetNullableString_FromString()
        Dim result As String = GetNullableString("test")
        Assert.AreEqual("test", result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToBool_FromBool()
        Dim result As Boolean = GetNullable(Of Boolean)(True)
        Assert.AreEqual(True, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToInteger_FromInteger()
        Dim result As Integer = GetNullable(Of Integer)(12)
        Assert.AreEqual(12, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToNullableBool_FromNull()
        Dim result As System.Nullable(Of Boolean) = GetNullable(Of System.Nullable(Of Boolean))(Nothing)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToNullableInteger_FromNull()
        Dim result As System.Nullable(Of Integer) = GetNullable(Of System.Nullable(Of Integer))(Nothing)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToNullableBool_FromDBNull()
        Dim result As System.Nullable(Of Boolean) = GetNullable(Of System.Nullable(Of Boolean))(DBNull.Value)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullable_ToNullableInteger_FromDBNull()
        Dim result As System.Nullable(Of Integer) = GetNullable(Of System.Nullable(Of Integer))(DBNull.Value)
        Assert.AreEqual(Nothing, result)
    End Sub

End Class
