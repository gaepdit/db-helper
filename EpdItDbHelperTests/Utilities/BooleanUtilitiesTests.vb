Imports EpdItDbHelper.DBUtilities

Namespace EpdItDbHelper.Tests

    <TestClass()> Public Class BooleanUtilitiesTests

        <TestMethod()>
        Public Sub ConvertDBValueToBoolean_GoodValues()
            Dim oneResult As Boolean = ConvertDBValueToBoolean("1", BooleanDBConversionType.OneOrZero)
            Assert.IsTrue(oneResult)

            Dim zeroResult As Boolean = ConvertDBValueToBoolean("0", BooleanDBConversionType.OneOrZero)
            Assert.IsFalse(zeroResult)

        End Sub

        <TestMethod()>
        <ExpectedException(GetType(FormatException))>
        Public Sub ConvertDBValueToBoolean_BadEmptyValue()
            Dim emptyStringResult As Boolean = ConvertDBValueToBoolean("", BooleanDBConversionType.OneOrZero)
        End Sub

        <TestMethod()>
        <ExpectedException(GetType(FormatException))>
        Public Sub ConvertDBValueToBoolean_BadNonEmptyValue()
            Dim emptyStringResult As Boolean = ConvertDBValueToBoolean("booger", BooleanDBConversionType.OneOrZero)
        End Sub

        <TestMethod()> Public Sub ConvertBooleanToDBValueTest()
            Dim oneResult As String = ConvertBooleanToDBValue(True, BooleanDBConversionType.OneOrZero)
            Assert.AreEqual("1", oneResult)

            Dim zeroResult As String = ConvertBooleanToDBValue(False, BooleanDBConversionType.OneOrZero)
            Assert.AreEqual("0", zeroResult)
        End Sub

        <TestMethod()> Public Sub StoreNothingIfZeroTest_Integer()
            Dim NonZeroInput As Integer = 13
            Dim NonZeroInputResult As Integer? = StoreNothingIfZero(NonZeroInput)
            Assert.AreEqual(NonZeroInput, NonZeroInputResult)

            Dim ZeroInput As Integer = 0
            Dim ZeroResult As Integer? = StoreNothingIfZero(ZeroInput)
            Assert.IsNull(ZeroResult)
            Assert.AreEqual(Nothing, ZeroResult)
        End Sub

        <TestMethod()> Public Sub StoreNothingIfZeroTest_Decimal()
            Dim NonZeroInput As Decimal = 13
            Dim NonZeroInputResult As Decimal? = StoreNothingIfZero(NonZeroInput)
            Assert.AreEqual(NonZeroInput, NonZeroInputResult)

            Dim ZeroInput As Decimal = 0
            Dim ZeroResult As Decimal? = StoreNothingIfZero(ZeroInput)
            Assert.IsNull(ZeroResult)
            Assert.AreEqual(Nothing, ZeroResult)
        End Sub
    End Class

End Namespace


