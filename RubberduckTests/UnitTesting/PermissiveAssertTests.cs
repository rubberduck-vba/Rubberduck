using System;
using System.Globalization;
using System.Reflection;
using NUnit.Framework;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [NonParallelizable]
    [Category("PermissiveAsserts")]
    public class PermissiveAssertTests
    {
        private AssertCompletedEventArgs _args;

        [SetUp]
        public void Initialize()
        {
            AssertHandler.OnAssertCompleted += AssertHandler_OnAssertCompleted;
        }

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [TearDown]
        public void Cleanup()
        {
            _args = null;
            AssertHandler.OnAssertCompleted -= AssertHandler_OnAssertCompleted;
        }

        // Note: nulls are basically equivalent to an empty variant. Therefore,
        // comparisons to a initialized variable should be as if it was the 
        // default value of the given type (e.g. 0 for numeric data types, empty
        // string for string data type and so on). To compare a value against a 
        // VBA's Null, we must use DBNull instead.

        [Test]
        [TestCase(true, -1)]
        [TestCase(true, "-1")]
        [TestCase(true, "true")]
        [TestCase(0, "0E1")]
        [TestCase(0e1, "0")]
        [TestCase("", null)]
        [TestCase(0, null)]
        [TestCase(null, 0)]
        [TestCase(null, "")]
        [TestCase("123", 123)]
        [TestCase(456, "456")]
        [TestCase("abc", "abc")]
        [TestCase("abc", "ABC")]
        public void PermissiveAreEqual(object left, object right)
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(left, right);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Test]
        [TestCase(true, 1)]
        [TestCase(true, 0)]
        [TestCase(true, "0")]
        [TestCase(true, "false")]
        [TestCase(123, "abc")]
        [TestCase("abc","def")]
        [TestCase("ABC", "def")]
        public void PermissiveAreNotEqual(object left, object right)
        {
            var assert = new PermissiveAssertClass();
            assert.AreNotEqual(left, right);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Test]
        public void PermissiveAreEqualStrings()
        {
            PermissiveAreEqual(string.Empty, "");
        }

        [Test]
        [TestCase("", false)]
        [TestCase("", true)]
        [TestCase(0, false)]
        [TestCase(0, true)]
        public void PermissiveAreNotEqualToNull(object x, bool invert)
        {
            if (invert)
            {
                PermissiveAreNotEqual(DBNull.Value, x);
            }
            else
            {
                PermissiveAreNotEqual(x, DBNull.Value);
            }
        }

        [Test]
        [TestCase("", false)]
        [TestCase("", true)]
        [TestCase(0, false)]
        [TestCase(0, true)]
        public void PermissiveAreNotEqualToMissing(object x, bool invert)
        {
            if (invert)
            {
                PermissiveAreNotEqual(Missing.Value, x);
            }
            else
            {
                PermissiveAreNotEqual(x, Missing.Value);
            }
        }

        [Test]
        [TestCase("0", "0", true)]
        [TestCase("1", "1", true)]
        [TestCase("1", "0", false)]
        [TestCase("1.1", "1.1", true)]
        [TestCase("1.1", "2.2", false)]
        public void PermissiveCompareDecimal(string x, string y, bool shouldEqual)
        {
            var dx = decimal.Parse(x, CultureInfo.InvariantCulture);
            var dy = decimal.Parse(y, CultureInfo.InvariantCulture);

            if (shouldEqual)
            {
                PermissiveAreEqual(dx, dy);
            }
            else
            {
                PermissiveAreNotEqual(dx, dy);
            }
        }
    }
}
