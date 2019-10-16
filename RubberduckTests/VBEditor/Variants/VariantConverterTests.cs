using System;
using System.Globalization;
using System.Runtime.InteropServices;
using NUnit.Framework;
using Rubberduck.VBEditor.Variants;

namespace RubberduckTests.VBEditor.Variants
{
    [TestFixture]
    public class VariantConverterTests
    {
        private const string TheOneTrueDateFormat = "yyyy-MM-dd HH:mm:ss";

        [Test]
        [TestCase(true, typeof(bool), ExpectedResult = true)]
        [TestCase(false, typeof(bool), ExpectedResult = false)]
        [TestCase("1", typeof(bool), ExpectedResult = true)]
        [TestCase("0", typeof(bool), ExpectedResult = false)]
        [TestCase("-1", typeof(bool), ExpectedResult = true)]
        [TestCase(1.0, typeof(bool), ExpectedResult = true)]
        [TestCase(0.0, typeof(bool), ExpectedResult = false)]
        [TestCase(-0.0, typeof(bool), ExpectedResult = false)]
        [TestCase(-1.0, typeof(bool), ExpectedResult = true)]
        [TestCase(1, typeof(bool), ExpectedResult = true)]
        [TestCase(0, typeof(bool), ExpectedResult = false)]
        [TestCase(-0, typeof(bool), ExpectedResult = false)]
        [TestCase(-1, typeof(bool), ExpectedResult = true)]

        [TestCase(true, typeof(byte), ExpectedResult = 255)]
        [TestCase(false, typeof(byte), ExpectedResult = 0)]
        [TestCase("1", typeof(byte), ExpectedResult = 1)]
        [TestCase("0", typeof(byte), ExpectedResult = 0)]
        [TestCase("255", typeof(byte), ExpectedResult = 255)]
        [TestCase(1.0, typeof(byte), ExpectedResult = 1)]
        [TestCase(0.0, typeof(byte), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(byte), ExpectedResult = 0)]
        [TestCase(255.0, typeof(byte), ExpectedResult = 255)]
        [TestCase(1, typeof(byte), ExpectedResult = 1)]
        [TestCase(0, typeof(byte), ExpectedResult = 0)]
        [TestCase(-0, typeof(byte), ExpectedResult = 0)]
        [TestCase(255, typeof(byte), ExpectedResult = 255)]

        [TestCase(true, typeof(sbyte), ExpectedResult = -1)]
        [TestCase(false, typeof(sbyte), ExpectedResult = 0)]
        [TestCase("1", typeof(sbyte), ExpectedResult = 1)]
        [TestCase("0", typeof(sbyte), ExpectedResult = 0)]
        [TestCase("-1", typeof(sbyte), ExpectedResult = -1)]
        [TestCase(1.0, typeof(sbyte), ExpectedResult = 1)]
        [TestCase(0.0, typeof(sbyte), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(sbyte), ExpectedResult = 0)]
        [TestCase(-1.0, typeof(sbyte), ExpectedResult = -1)]
        [TestCase(1, typeof(sbyte), ExpectedResult = 1)]
        [TestCase(0, typeof(sbyte), ExpectedResult = 0)]
        [TestCase(-0, typeof(sbyte), ExpectedResult = 0)]
        [TestCase(-1, typeof(sbyte), ExpectedResult = -1)]

        [TestCase(true, typeof(short), ExpectedResult = -1)]
        [TestCase(false, typeof(short), ExpectedResult = 0)]
        [TestCase("1", typeof(short), ExpectedResult = 1)]
        [TestCase("0", typeof(short), ExpectedResult = 0)]
        [TestCase("-1", typeof(short), ExpectedResult = -1)]
        [TestCase(1.0, typeof(short), ExpectedResult = 1)]
        [TestCase(0.0, typeof(short), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(short), ExpectedResult = 0)]
        [TestCase(-1.0, typeof(short), ExpectedResult = -1)]
        [TestCase(1, typeof(short), ExpectedResult = 1)]
        [TestCase(0, typeof(short), ExpectedResult = 0)]
        [TestCase(-0, typeof(short), ExpectedResult = 0)]
        [TestCase(-1, typeof(short), ExpectedResult = -1)]

        [TestCase(true, typeof(ushort), ExpectedResult = 65535)]
        [TestCase(false, typeof(ushort), ExpectedResult = 0)]
        [TestCase("1", typeof(ushort), ExpectedResult = 1)]
        [TestCase("0", typeof(ushort), ExpectedResult = 0)]
        [TestCase("255", typeof(ushort), ExpectedResult = 255)]
        [TestCase(1.0, typeof(ushort), ExpectedResult = 1)]
        [TestCase(0.0, typeof(ushort), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(ushort), ExpectedResult = 0)]
        [TestCase(255.0, typeof(ushort), ExpectedResult = 255)]
        [TestCase(1, typeof(ushort), ExpectedResult = 1)]
        [TestCase(0, typeof(ushort), ExpectedResult = 0)]
        [TestCase(-0, typeof(ushort), ExpectedResult = 0)]
        [TestCase(255, typeof(ushort), ExpectedResult = 255)]

        [TestCase(true, typeof(int), ExpectedResult = -1)]
        [TestCase(false, typeof(int), ExpectedResult = 0)]
        [TestCase("1", typeof(int), ExpectedResult = 1)]
        [TestCase("0", typeof(int), ExpectedResult = 0)]
        [TestCase("-1", typeof(int), ExpectedResult = -1)]
        [TestCase(1.0, typeof(int), ExpectedResult = 1)]
        [TestCase(0.0, typeof(int), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(int), ExpectedResult = 0)]
        [TestCase(-1.0, typeof(int), ExpectedResult = -1)]
        [TestCase(1, typeof(int), ExpectedResult = 1)]
        [TestCase(0, typeof(int), ExpectedResult = 0)]
        [TestCase(-0, typeof(int), ExpectedResult = 0)]
        [TestCase(-1, typeof(int), ExpectedResult = -1)]

        [TestCase(true, typeof(uint), ExpectedResult = 4294967295)]
        [TestCase(false, typeof(uint), ExpectedResult = 0)]
        [TestCase("1", typeof(uint), ExpectedResult = 1)]
        [TestCase("0", typeof(uint), ExpectedResult = 0)]
        [TestCase("255", typeof(uint), ExpectedResult = 255)]
        [TestCase(1.0, typeof(uint), ExpectedResult = 1)]
        [TestCase(0.0, typeof(uint), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(uint), ExpectedResult = 0)]
        [TestCase(255.0, typeof(uint), ExpectedResult = 255)]
        [TestCase(1, typeof(uint), ExpectedResult = 1)]
        [TestCase(0, typeof(uint), ExpectedResult = 0)]
        [TestCase(-0, typeof(uint), ExpectedResult = 0)]
        [TestCase(255, typeof(uint), ExpectedResult = 255)]

        [TestCase(true, typeof(long), ExpectedResult = -1)]
        [TestCase(false, typeof(long), ExpectedResult = 0)]
        [TestCase("1", typeof(long), ExpectedResult = 1)]
        [TestCase("0", typeof(long), ExpectedResult = 0)]
        [TestCase("-1", typeof(long), ExpectedResult = -1)]
        [TestCase(1.0, typeof(long), ExpectedResult = 1)]
        [TestCase(0.0, typeof(long), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(long), ExpectedResult = 0)]
        [TestCase(-1.0, typeof(long), ExpectedResult = -1)]
        [TestCase(1, typeof(long), ExpectedResult = 1)]
        [TestCase(0, typeof(long), ExpectedResult = 0)]
        [TestCase(-0, typeof(long), ExpectedResult = 0)]
        [TestCase(-1, typeof(long), ExpectedResult = -1)]

        [TestCase(true, typeof(ulong), ExpectedResult = 18446744073709551615)]
        [TestCase(false, typeof(ulong), ExpectedResult = 0)]
        [TestCase("1", typeof(ulong), ExpectedResult = 1)]
        [TestCase("0", typeof(ulong), ExpectedResult = 0)]
        [TestCase("255", typeof(ulong), ExpectedResult = 255)]
        [TestCase(1.0, typeof(ulong), ExpectedResult = 1)]
        [TestCase(0.0, typeof(ulong), ExpectedResult = 0)]
        [TestCase(-0.0, typeof(ulong), ExpectedResult = 0)]
        [TestCase(255.0, typeof(ulong), ExpectedResult = 255)]
        [TestCase(1, typeof(ulong), ExpectedResult = 1)]
        [TestCase(0, typeof(ulong), ExpectedResult = 0)]
        [TestCase(-0, typeof(ulong), ExpectedResult = 0)]
        [TestCase(255, typeof(ulong), ExpectedResult = 255)]

        [TestCase(true, typeof(float), ExpectedResult = -1f)]
        [TestCase(false, typeof(float), ExpectedResult = 0f)]
        [TestCase("1", typeof(float), ExpectedResult = 1f)]
        [TestCase("0", typeof(float), ExpectedResult = 0f)]
        [TestCase("-1", typeof(float), ExpectedResult = -1f)]
        [TestCase(1.0, typeof(float), ExpectedResult = 1.0f)]
        [TestCase(0.0, typeof(float), ExpectedResult = 0.0f)]
        [TestCase(-0.0, typeof(float), ExpectedResult = -0.0f)]
        [TestCase(-1.0, typeof(float), ExpectedResult = -1.0f)]
        [TestCase(1, typeof(float), ExpectedResult = 1f)]
        [TestCase(0, typeof(float), ExpectedResult = 0f)]
        [TestCase(-0, typeof(float), ExpectedResult = 0f)]
        [TestCase(-1, typeof(float), ExpectedResult = -1f)]

        [TestCase(true, typeof(double), ExpectedResult = -1d)]
        [TestCase(false, typeof(double), ExpectedResult = 0d)]
        [TestCase("1", typeof(double), ExpectedResult = 1d)]
        [TestCase("0", typeof(double), ExpectedResult = 0d)]
        [TestCase("-1", typeof(double), ExpectedResult = -1d)]
        [TestCase(1.0, typeof(double), ExpectedResult = 1.0d)]
        [TestCase(0.0, typeof(double), ExpectedResult = 0.0d)]
        [TestCase(-0.0, typeof(double), ExpectedResult = -0.0d)]
        [TestCase(-1.0, typeof(double), ExpectedResult = -1.0d)]
        [TestCase(1, typeof(double), ExpectedResult = 1d)]
        [TestCase(0, typeof(double), ExpectedResult = 0d)]
        [TestCase(-0, typeof(double), ExpectedResult = 0d)]
        [TestCase(-1, typeof(double), ExpectedResult = -1d)]

        [TestCase(true, typeof(decimal), ExpectedResult = -1)]
        [TestCase(false, typeof(decimal), ExpectedResult = 0)]
        [TestCase("1", typeof(decimal), ExpectedResult = 1)]
        [TestCase("0", typeof(decimal), ExpectedResult = 0)]
        [TestCase("-1", typeof(decimal), ExpectedResult = -1)]
        [TestCase(1.0, typeof(decimal), ExpectedResult = 1.0)]
        [TestCase(0.0, typeof(decimal), ExpectedResult = 0.0)]
        [TestCase(-0.0, typeof(decimal), ExpectedResult = -0.0)]
        [TestCase(-1.0, typeof(decimal), ExpectedResult = -1.0)]
        [TestCase(1, typeof(decimal), ExpectedResult = 1)]
        [TestCase(0, typeof(decimal), ExpectedResult = 0)]
        [TestCase(-0, typeof(decimal), ExpectedResult = 0)]
        [TestCase(-1, typeof(decimal), ExpectedResult = -1)]
        [TestCase(1.0d, typeof(string), ExpectedResult = "1")]
        [TestCase(0.0d, typeof(string), ExpectedResult = "0")]
        // Unstable test case - it will pass in VS runnner but fail in Resharper runner
        // [TestCase(-0.0d, typeof(string), ExpectedResult = "0")]
        [TestCase(-1.0d, typeof(string), ExpectedResult = "-1")]
        [TestCase(true, typeof(string), ExpectedResult = "-1")]
        [TestCase(false, typeof(string), ExpectedResult = "0")]
        [TestCase("1", typeof(string), ExpectedResult = "1")]
        [TestCase("0", typeof(string), ExpectedResult = "0")]
        [TestCase("-1", typeof(string), ExpectedResult = "-1")]
        [TestCase(1, typeof(string), ExpectedResult = "1")]
        [TestCase(0, typeof(string), ExpectedResult = "0")]
        [TestCase(-0, typeof(string), ExpectedResult = "0")]
        [TestCase(-1, typeof(string), ExpectedResult = "-1")]

        public object Test_ObjectConversion_SimpleValues(object value, Type targetType)
        {
            var result = VariantConverter.ChangeType(value, targetType);

            if (result is DateTime dt)
            {
                return dt.ToString(TheOneTrueDateFormat);
            }

            return result;
        }

        /// <remarks>
        /// This assumes en-US locale. 
        /// See <see cref="Test_US_format_String_To_Date_Localized"/> for localization behavior
        /// </remarks>
        [Test]
        [TestCase(true, typeof(DateTime), ExpectedResult = "1899-12-29 00:00:00")]
        [TestCase(false, typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase("1899/12/31", typeof(DateTime), ExpectedResult = "1899-12-31 00:00:00")]
        [TestCase("1899/12/30", typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase("1899/12/29", typeof(DateTime), ExpectedResult = "1899-12-29 00:00:00")]
        [TestCase("1899-12-30", typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase("04/12/2000", typeof(DateTime), ExpectedResult = "2000-04-12 00:00:00")]
        [TestCase("12/04/2000", typeof(DateTime), ExpectedResult = "2000-12-04 00:00:00")]
        [TestCase("13/04/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("04/13/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase(1.0, typeof(DateTime), ExpectedResult = "1899-12-31 00:00:00")]
        [TestCase(0.0, typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase(-0.0, typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase(-1.0, typeof(DateTime), ExpectedResult = "1899-12-29 00:00:00")]
        [TestCase(1, typeof(DateTime), ExpectedResult = "1899-12-31 00:00:00")]
        [TestCase(0, typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase(-0, typeof(DateTime), ExpectedResult = "1899-12-30 00:00:00")]
        [TestCase(-1, typeof(DateTime), ExpectedResult = "1899-12-29 00:00:00")]
        [TestCase(-657434.00001157413d, typeof(DateTime), ExpectedResult = "0100-01-01 00:00:01")]
        [TestCase(-657434.000011574d, typeof(DateTime), ExpectedResult = "0100-01-01 00:00:01")]
        [TestCase(2958465.999988426d, typeof(DateTime), ExpectedResult = "9999-12-31 23:59:59")]
        [TestCase(2958465.99998843d, typeof(DateTime), ExpectedResult = "9999-12-31 23:59:59")]
        [TestCase(0.5d, typeof(DateTime), ExpectedResult = "1899-12-30 12:00:00")]
        [TestCase(-0.5d, typeof(DateTime), ExpectedResult = "1899-12-30 12:00:00")]
        public string Test_ObjectConversion_SimpleValues_To_Date(object value, Type targetType)
        {
            var culture = new CultureInfo("en-US");
            var result = VariantConverter.ChangeType(value, targetType, culture);
            if (result is DateTime dt)
            {
                return dt.ToString(TheOneTrueDateFormat);
            }

            // Invalid result
            return string.Empty;
        }

        /// <remarks>
        /// Note that de-DE results differ from fr-CA and en-US
        /// </remarks>
        [TestCase("fr-CA", "04/12/2000", typeof(DateTime), ExpectedResult = "2000-04-12 00:00:00")]
        [TestCase("fr-CA", "12/04/2000", typeof(DateTime), ExpectedResult = "2000-12-04 00:00:00")]
        [TestCase("fr-CA", "13/04/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("fr-CA", "04/13/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("de-DE", "04/12/2000", typeof(DateTime), ExpectedResult = "2000-12-04 00:00:00")]
        [TestCase("de-DE", "12/04/2000", typeof(DateTime), ExpectedResult = "2000-04-12 00:00:00")]
        [TestCase("de-DE", "13/04/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("de-DE", "04/13/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("en-US", "04/12/2000", typeof(DateTime), ExpectedResult = "2000-04-12 00:00:00")]
        [TestCase("en-US", "12/04/2000", typeof(DateTime), ExpectedResult = "2000-12-04 00:00:00")]
        [TestCase("en-US", "13/04/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        [TestCase("en-US", "04/13/2000", typeof(DateTime), ExpectedResult = "2000-04-13 00:00:00")]
        public string Test_US_format_String_To_Date_Localized(string locale, object value, Type targetType)
        {
            var culture = new CultureInfo(locale);
            var result = VariantConverter.ChangeType(value, targetType, culture);
            if (result is DateTime dt)
            {
                return dt.ToString(TheOneTrueDateFormat);
            }

            // Invalid result
            return string.Empty;
        }

        /// <remarks>
        /// Do NOT write localization-sensitive tests in this test.
        /// Use <see cref="Test_ObjectConversion_Currency_Localized"/> instead
        /// </remarks>
        [Test]
        [TestCase(1.2, typeof(int), ExpectedResult = 1)]
        [TestCase(1.2, typeof(float), ExpectedResult = 1.2f)]
        [TestCase(1.2, typeof(double), ExpectedResult = 1.2d)]
        [TestCase(1.2, typeof(decimal), ExpectedResult = 1.2)]
        [TestCase(1.2, typeof(DateTime), ExpectedResult = "1899-12-31 04:48:00")]
        public object Test_ObjectConversion_Currency(decimal value, Type targetType)
        {
            var cy = new CurrencyWrapper(value);
            var result = VariantConverter.ChangeType(cy, targetType);

            if (result is DateTime dt)
            {
                return dt.ToString(TheOneTrueDateFormat);
            }

            return result;
        }

        [Test]
        [TestCase(1.0, typeof(string), "en-US", ExpectedResult = "1")]
        [TestCase(0.0, typeof(string), "en-US", ExpectedResult = "0")]
        [TestCase(-1.0, typeof(string), "en-US", ExpectedResult = "-1")]
        [TestCase(0.1, typeof(string), "en-US", ExpectedResult = "0.1")]
        [TestCase(-0.1, typeof(string), "en-US", ExpectedResult = "-0.1")]
        [TestCase(1.0, typeof(string), "de-DE", ExpectedResult = "1")]
        [TestCase(0.0, typeof(string), "de-DE", ExpectedResult = "0")]
        [TestCase(-1.0, typeof(string), "de-DE", ExpectedResult = "-1")]
        [TestCase(0.1, typeof(string), "de-DE", ExpectedResult = "0,1")]
        [TestCase(-0.1, typeof(string), "de-DE", ExpectedResult = "-0,1")]
        public object Test_ObjectConversion_Currency_Localized(decimal value, Type targetType, string locale)
        {
            var cy = new CurrencyWrapper(value);
            return VariantConverter.ChangeType(cy, targetType, new CultureInfo(locale));
        }

        /// <remarks>
        /// Do NOT write localization-sensitive tests in this test.
        /// Use <see cref="Test_ObjectConversion_Date_Localized"/> instead
        ///
        /// Also, some double values can be ambiguous. See <see cref="Test_ObjectConversion_SimpleValues_To_Date"/>
        /// for the coverages of those ambiguous results.
        /// </remarks>
        [Test]
        [TestCase("1899-12-30 00:00:00", typeof(int), ExpectedResult = 0)]
        [TestCase("1899-12-31 00:00:00", typeof(int), ExpectedResult = 1)]
        [TestCase("1899-12-29 00:00:00", typeof(int), ExpectedResult = -1)]
        [TestCase("1899-12-30 00:00:00", typeof(double), ExpectedResult = 0d)]
        [TestCase("1899-12-31 00:00:00", typeof(double), ExpectedResult = 1d)]
        [TestCase("1899-12-29 00:00:00", typeof(double), ExpectedResult = -1d)]
        [TestCase("1899-12-30 12:00:00", typeof(double), ExpectedResult = 0.5d)]
        [TestCase("1899-12-29 12:00:00", typeof(double), ExpectedResult = -1.5d)]
        [TestCase("0100-01-01 00:00:00", typeof(double), ExpectedResult = -657434d)]
        [TestCase("0100-01-01 00:00:01", typeof(double), ExpectedResult = -657434.00001157413d)]    // Note: VBA returns -657434.000011574d but accepts the other values as equal
        [TestCase("9999-12-31 23:59:59", typeof(double), ExpectedResult = 2958465.999988426d)]      // Note: VBA returns 2958465.99998843d but accepts the other values as equal
        public object Test_ObjectConversion_Date(string value, Type targetType)
        {
            var date = DateTime.ParseExact(value, TheOneTrueDateFormat, CultureInfo.InvariantCulture);
            return VariantConverter.ChangeType(date, targetType);
        }

        [Test]
        [TestCase("1899-12-30 00:00:00", typeof(string), "en-US", ExpectedResult = "12:00:00 AM")]
        [TestCase("1899-12-31 00:00:00", typeof(string), "en-US", ExpectedResult = "12/31/1899")]
        [TestCase("1899-12-29 00:00:00", typeof(string), "en-US", ExpectedResult = "12/29/1899")]
        [TestCase("1899-12-30 01:00:00", typeof(string), "en-US", ExpectedResult = "1:00:00 AM")]
        [TestCase("1899-12-31 11:00:00", typeof(string), "en-US", ExpectedResult = "12/31/1899 11:00:00 AM")]
        [TestCase("1899-12-29 13:00:00", typeof(string), "de-DE", ExpectedResult = "29.12.1899 13:00:00")]
        [TestCase("1899-12-30 00:00:00", typeof(string), "de-DE", ExpectedResult = "00:00:00")]
        [TestCase("1899-12-31 00:00:00", typeof(string), "de-DE", ExpectedResult = "31.12.1899")]
        [TestCase("1899-12-29 00:00:00", typeof(string), "de-DE", ExpectedResult = "29.12.1899")]
        [TestCase("1899-12-30 01:00:00", typeof(string), "de-DE", ExpectedResult = "01:00:00")]
        [TestCase("1899-12-31 11:00:00", typeof(string), "de-DE", ExpectedResult = "31.12.1899 11:00:00")]
        [TestCase("1899-12-29 13:00:00", typeof(string), "de-DE", ExpectedResult = "29.12.1899 13:00:00")]
        public object Test_ObjectConversion_Date_Localized(string value, Type targetType, string locale)
        {
            var culture = new CultureInfo(locale);
            var date = DateTime.ParseExact(value, TheOneTrueDateFormat, culture);
            return VariantConverter.ChangeType(date, targetType, culture);
        }

        [Test]
        [TestCase(typeof(string))]
        [TestCase(typeof(byte))]
        [TestCase(typeof(short))]
        [TestCase(typeof(ushort))]
        [TestCase(typeof(int))]
        [TestCase(typeof(uint))] // the doc says default marshal will use it but this doesn't seem to be the case when converting
        [TestCase(typeof(long))]
        [TestCase(typeof(ulong))]
        public void Test_Error_Is_Not_Convertible(Type targetType)
        {
            Assert.Throws<COMException>(() =>
            {
                var err = new ErrorWrapper(1);
                VariantConverter.ChangeType(err, targetType);
            });
        }

        [Test]
        public void Test_ObjectConversion_Object()
        {
            var obj = new object();
            var unk = new UnknownWrapper(obj);
            var result = VariantConverter.ChangeType(unk, typeof(DispatchWrapper));

            Assert.AreSame(obj, result);

            var result2 = VariantConverter.ChangeType(result, typeof(UnknownWrapper));

            Assert.AreSame(unk.WrappedObject, result2);
        }

        [Test]
        [TestCase(1)]
        [TestCase("1")]
        [TestCase(1.0)]
        [TestCase("")]
        [TestCase(null)]
        public void Test_ObjectConversion_DbNull(object value)
        {
            var result = VariantConverter.ChangeType(value, VARENUM.VT_NULL);
            Assert.IsInstanceOf<DBNull>(result);
        }

        [Test]
        [TestCase(1)]
        [TestCase("1")]
        [TestCase(1.0)]
        [TestCase("")]
        [TestCase(null)]
        public void Test_ObjectConversion_Null(object value)
        {
            var result = VariantConverter.ChangeType(value, VARENUM.VT_EMPTY);
            Assert.IsNull(result);
        }

        [Test]
        [TestCase(VARENUM.VT_NULL, TypeCode.DBNull, ExpectedResult = typeof(DBNull))]
        [TestCase(VARENUM.VT_BSTR, TypeCode.String, ExpectedResult = typeof(string))]
        [TestCase(VARENUM.VT_CY, TypeCode.Decimal, ExpectedResult = typeof(decimal))]
        [TestCase(VARENUM.VT_DATE, TypeCode.DateTime, ExpectedResult = typeof(DateTime))]
        [TestCase(VARENUM.VT_DECIMAL, TypeCode.Decimal, ExpectedResult = typeof(decimal))]
        [TestCase(VARENUM.VT_I2, TypeCode.Int16, ExpectedResult = typeof(short))]
        [TestCase(VARENUM.VT_I4, TypeCode.Int32, ExpectedResult = typeof(int))]
        [TestCase(VARENUM.VT_I8, TypeCode.Int64, ExpectedResult = typeof(long))]
        [TestCase(VARENUM.VT_R4, TypeCode.Single, ExpectedResult = typeof(float))]
        [TestCase(VARENUM.VT_R8, TypeCode.Double, ExpectedResult = typeof(double))]
        [TestCase(VARENUM.VT_UI2, TypeCode.UInt16, ExpectedResult = typeof(ushort))]
        [TestCase(VARENUM.VT_UI4, TypeCode.UInt32, ExpectedResult = typeof(uint))]
        [TestCase(VARENUM.VT_UI8, TypeCode.UInt64, ExpectedResult = typeof(ulong))]
        [TestCase(VARENUM.VT_UNKNOWN, TypeCode.Object, ExpectedResult = typeof(ConvertibleTest))]
        public Type Test_ObjectConversion_IConvertible(VARENUM vt, TypeCode code)
        {
            var convertible = new ConvertibleTest(code);
            var result = VariantConverter.ChangeType(convertible, vt);
            return result.GetType();
        }

        [Test]
        public void Test_ObjectConversion_IConvertible_Null()
        {
            var convertible = new ConvertibleTest(TypeCode.Empty);
            var result = VariantConverter.ChangeType(convertible, VARENUM.VT_EMPTY);
            Assert.IsNull(result);
        }
    }
}