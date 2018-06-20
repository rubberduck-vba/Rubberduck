using NUnit.Framework;
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
        public void AreNotEqualShouldBeInconclusiveWithDifferentTypes()
        {
            int obj1 = 10;
            double obj2 = 10;

            var assert = new PermissiveAssertClass();
            assert.AreNotEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
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
        public void NullIsNotComparableWithValues()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void DifferentTypesEqualitySucceeds()
        {
            var assert = new PermissiveAssertClass();
            assert.AreEqual(42, "42");

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void OnAssertSucceeded_ReturnsResultSuccess()
        {
            AssertHandler.OnAssertSucceeded();

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }
    }
}
