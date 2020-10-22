using System;
using NUnit.Framework;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    public class AssertTests
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

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [Category("Unit Testing")]
        [Test]
        public void IsTrueSucceedsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(true);
            
            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsTrueFailsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(false);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsFalseSucceedsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(false);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsFalseFailsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(true);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldSucceedWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = GetComObject();
            var obj2 = obj1;
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldFailWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = GetComObject();
            var obj2 = GetComObject();
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldSucceedWithTwoNullReferences()
        {
            var assert = new AssertClass();
            assert.AreSame(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldFailWithActualNullReference()
        {
            var assert = new AssertClass();
            assert.AreSame(GetComObject(), null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreSameShouldFailWithExpectedNullReference()
        {
            var assert = new AssertClass();
            assert.AreSame(null, GetComObject());

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotSameShouldSucceedWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = GetComObject();
            var obj2 = GetComObject();
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotSameShouldSuccedWithOneNullReference()
        {
            var assert = new AssertClass();
            assert.AreNotSame(GetComObject(), null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotSameShouldFailWithBothNullReferences()
        {
            var assert = new AssertClass();
            assert.AreNotSame(null, null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotSameShouldFailWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = GetComObject();
            var obj2 = obj1;
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldSucceedWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 1);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreEqualShouldFailWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotEqualShouldSucceedWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotEqualShouldFailWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 1);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void AreNotEqualShouldBeInconclusiveWithDifferentTypes()
        {
            int obj1 = 10;
            double obj2 = 10;

            var assert = new AssertClass();
            assert.AreNotEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNothingShouldSucceedWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNothingShouldFailWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(new object());

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNotNothingShouldFailWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void IsNotNothingShouldSucceedWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(new object());

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void FailShouldFail()
        {
            var assert = new AssertClass();
            assert.Fail();

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void InconclusiveShouldBeInconclusive()
        {
            var assert = new AssertClass();
            assert.Inconclusive();

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void NullValuesAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void NullAndEmptyStringAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, string.Empty);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void NullIsNotComparableWithValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void DifferentTypesEqualityIsNotConclusive()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, "42");

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void OnAssertSucceeded_ReturnsResultSuccess()
        {
            AssertHandler.OnAssertSucceeded();

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void OnAssertFailed_ReturnsResultFailed()
        {
            AssertHandler.OnAssertFailed("MyMethod", "I Failed");

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [Category("Unit Testing")]
        [Test]
        public void OnAssertInconclusive_ReturnsResultInconclusive()
        {
            AssertHandler.OnAssertInconclusive("Inconclusive");

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        private static Type GetComObjectType() => Type.GetTypeFromProgID("Scripting.FileSystemObject");
        private object GetComObject() => Activator.CreateInstance(GetComObjectType());
    }
}
