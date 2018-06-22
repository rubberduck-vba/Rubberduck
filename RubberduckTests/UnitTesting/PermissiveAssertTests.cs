using System;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    //NOTE: The tests for reference equity are commented out pending some way of figuring out how to test the correct behavior.
    //These methods have to check to see if the parameters are COM objects (see https://github.com/rubberduck-vba/Rubberduck/issues/2848)
    //to make the result match the VBA interpretations of reference and value types.  Similarly, the SequenceEqual and NotSequenceEqual
    //methods remain untested because they make several of the same Type tests that are AFAIK impossible to mock.

    [TestFixture]
    public class PermissiveAssertTests
    {
        private AssertCompletedEventArgs _args;

        [SetUp]
        public void Initialize()
        {
            AssertHandler.OnAssertCompleted += AssertHandler_OnAssertCompleted;
        }

        [TearDown]
        public void Cleanup()
        {
            _args = null;
            AssertHandler.OnAssertCompleted -= AssertHandler_OnAssertCompleted;
        }

        [Category("Unit Testing")]
        [Test]
        public void IsTrueSucceedsWithTrueExpression()
        {
            var assert = new PermissiveAssertClass();
            assert.IsTrue(true);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsTrueFailsWithFalseExpression()
        {
            var assert = new PermissiveAssertClass();
            assert.IsTrue(false);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsFalseSucceedsWithFalseExpression()
        {
            var assert = new PermissiveAssertClass();
            assert.IsFalse(false);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsFalseFailsWithTrueExpression()
        {
            var assert = new PermissiveAssertClass();
            assert.IsFalse(true);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        //[Category("Unit Testing")]
        //[Test]
        //public void AreSameShouldSucceedWithSameReferences()
        //{
        //    var assert = new PermissiveAssertClass();
        //    var obj1 = new object();
        //    var obj2 = obj1;
        //    assert.AreSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        //[Category("Unit Testing")]
        //[Test]
        //public void AreSameShouldFailWithDifferentReferences()
        //{
        //    var assert = new PermissiveAssertClass();
        //    var obj1 = new object();
        //    var obj2 = new object();
        //    assert.AreSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldSucceedWithTwoNullReferences()
        {
            var assert = new PermissiveAssertClass();
            assert.AreSame(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        //[Category("Unit Testing")]
        //[Test]
        //public void AreSameShouldFailWithActualNullReference()
        //{
        //    var assert = new PermissiveAssertClass();
        //    assert.AreSame(new object(), null);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        //[Category("Unit Testing")]
        //[Test]
        //public void AreSameShouldFailWithExpectedNullReference()
        //{
        //    var assert = new PermissiveAssertClass();
        //    assert.AreSame(null, new object());

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        //[Category("Unit Testing")]
        //[Test]
        //public void AreNotSameShouldSucceedWithDifferentReferences()
        //{
        //    var assert = new PermissiveAssertClass();
        //    var obj1 = new object();
        //    var obj2 = new object();
        //    assert.AreNotSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        //[Category("Unit Testing")]
        //[Test]
        //public void AreNotSameShouldSuccedWithOneNullReference()
        //{
        //    var assert = new PermissiveAssertClass();
        //    assert.AreNotSame(new object(), null);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        [Category("Unit Testing")]
        [Test]
        public void AreNotSameShouldFailWithBothNullReferences()
        {
            var assert = new PermissiveAssertClass();
            assert.AreNotSame(null, null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        //[Category("Unit Testing")]
        //[Test]
        //public void AreNotSameShouldFailWithSameReferences()
        //{
        //    var assert = new PermissiveAssertClass();
        //    var obj1 = new object();
        //    var obj2 = obj1;
        //    assert.AreNotSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithSameValues()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(1, 1);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldFailWithDifferentValues()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(1, 2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotEqualShouldSucceedWithDifferentValues()
        {
            var assert = new PermissiveAssertClass();
            assert.AreNotEqual(1, 2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotEqualShouldFailWithSameValues()
        {
            var assert = new PermissiveAssertClass();
            assert.AreNotEqual(1, 1);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }


        [Category("Unit Testing")]
        [Test]
        public void IsNothingShouldSucceedWithNullValue()
        {
            var assert = new PermissiveAssertClass();
            assert.IsNothing(null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNothingShouldFailWithNonNullValue()
        {
            var assert = new PermissiveAssertClass();
            assert.IsNothing(new object());

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNotNothingShouldFailWithNullValue()
        {
            var assert = new PermissiveAssertClass();
            assert.IsNotNothing(null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNotNothingShouldSucceedWithNonNullValue()
        {
            var assert = new PermissiveAssertClass();
            assert.IsNotNothing(new object());

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void FailShouldFail()
        {
            var assert = new PermissiveAssertClass();
            assert.Fail();

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void InconclusiveShouldBeInconclusive()
        {
            var assert = new PermissiveAssertClass();
            assert.Inconclusive();

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [Category("Unit Testing")]
        [Test]
        public void NullValuesAreEqual()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void NullAndEmptyStringAreEqual()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(null, string.Empty);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void NullIsNotComparableWithNumbers()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void OnAssertSucceeded_ReturnsResultSuccess()
        {
            AssertHandler.OnAssertSucceeded();

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIdenticalStrings()
        {
            string obj1 = "foo";
            string obj2 = "foo";

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldFailWithDifferentStrings()
        {
            string obj1 = "foo";
            string obj2 = "bar";

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndInteger()
        {
            Int16 vbaInteger1 = 1;
            Int16 vbaInteger2 = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger1, vbaInteger2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndLong()
        {
            Int16 vbaInteger = 1;
            int vbaLong = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, vbaLong);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndString()
        {
            Int16 obj1 = 1;
            string obj2 = "1";

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndDouble()
        {
            Int16 vbaInteger = 1;
            double obj2 = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndByte()
        {
            Int16 vbaInteger = 1;
            byte vbaByte = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, vbaByte);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithIntegerAndSingle()
        {
            Int16 vbaInteger = 1;
            Single obj2 = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithLongAndLong()
        {
            Int16 vbaInteger1 = 1;
            Int16 vbaInteger2 = 1;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger1, vbaInteger2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("UnitTesting")]
        [Test]
        public void AreEqualShouldSucceedWithLongAndString()
        {
            Int16 vbaInteger = 1;
            string obj2 = "1";

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }


        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithLongAndDouble()
        {
            Int16 vbaInteger = 10;
            double obj2 = 10;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithLongAndByte()
        {
            Int16 vbaInteger = 10;
            byte obj2 = 10;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithLongAndSingle()
        {
            Int16 vbaInteger = 10;
            Single obj2 = 10;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(vbaInteger, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithStringAndString()
        {
            string string1 = "10";
            string string2 = "10";

            var assert = new PermissiveAssertClass();
            assert.AreEqual(string1, string2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithStringAndDouble()
        {
            string obj1 = "10.57";
            double obj2 = 10.57;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithStringAndByte()
        {
            string obj1 = "10";
            byte obj2 = 10;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithStringAndSingle()
        {
            string obj1 = "10.23";
            const Single obj2 = 10.23f;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithDoubleAndDouble()
        {
            double obj1 = 11.25;
            double obj2 = 11.25;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithDoubleAndByte()
        {
            double obj1 = 11;
            byte obj2 = 11;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithDoubleAndSingle()
        {
            double obj1 = 11.43;
            Single obj2 = 11.43f;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithByteAndByte()
        {
            byte obj1 = 11;
            byte obj2 = 11;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithByteAndSingle()
        {
            byte obj1 = 11;
            Single obj2 = 11;

            var assert = new PermissiveAssertClass();
            assert.AreEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }
    }
}
