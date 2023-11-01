Imports GaEpd.DBUtilities
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class NullableDateTimeUtilitiesTests

    <TestMethod>
    Public Sub GetNullableDateTime_FromValidString1()
        Dim result As Date? = GetNullableDateTime("2016-06-07")
        Assert.AreEqual(New Date(2016, 6, 7), result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromValidString2()
        Dim result As Date? = GetNullableDateTime("June 7, 2016")
        Assert.AreEqual(New Date(2016, 6, 7), result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromValidString3()
        Dim result As Date? = GetNullableDateTime("2016-06-07 08:00:00")
        Assert.AreEqual(New Date(2016, 6, 7, 8, 0, 0), result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromValidString4()
        Dim result As Date? = GetNullableDateTime("2016-06-07 14:00:00")
        Assert.AreEqual(New Date(2016, 6, 7, 14, 0, 0), result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromValidString5()
        Dim result As Date? = GetNullableDateTime("2016-06-07 2:00 pm")
        Assert.AreEqual(New Date(2016, 6, 7, 14, 0, 0), result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromDBNull()
        Dim result As Date? = GetNullableDateTime(DBNull.Value)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromNull()
        Dim result As Date? = GetNullableDateTime(Nothing)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromInvalidString()
        Dim result As Date? = GetNullableDateTime("test")
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromBool_True()
        Dim result As Date? = GetNullableDateTime(True)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromBool_False()
        Dim result As Date? = GetNullableDateTime(False)
        Assert.AreEqual(Nothing, result)
    End Sub

    <TestMethod>
    Public Sub GetNullableDateTime_FromInteger()
        Dim result As Date? = GetNullableDateTime(12)
        Assert.AreEqual(Nothing, result)
    End Sub

End Class
