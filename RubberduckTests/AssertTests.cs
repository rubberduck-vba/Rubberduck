using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests
{
    [TestClass]
    public class AssertTests
    {
        private AssertCompletedEventArgs _args;

        [TestInitialize]
        public void Initialize()
        {
            AssertHandler.OnAssertCompleted += AssertHandler_OnAssertCompleted;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _args = null;
            AssertHandler.OnAssertCompleted -= AssertHandler_OnAssertCompleted;
        }

        [TestMethod]
        public void IsTrueSucceedsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(true);
            
            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void IsTrueFailsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(false);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void IsFalseSucceedsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(false);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void IsFalseFailsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(true);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void AreSameShouldSucceedWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = obj1;
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void AreSameShouldFailWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = new object();
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void AreNotSameShouldSucceedWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = new object();
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void AreNotSameShouldFailWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = obj1;
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void AreEqualShouldSucceedWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 1);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void AreEqualShouldFailWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void AreNotEqualShouldSucceedWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 2);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void AreNotEqualShouldFailWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 1);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void IsNothingShouldSucceedWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(null);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void IsNothingShouldFailWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(new object());

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void IsNotNothingShouldFailWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(null);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void IsNotNothingShouldSucceedWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(new object());

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void FailShouldFail()
        {
            var assert = new AssertClass();
            assert.Fail();

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Failed);
        }

        [TestMethod]
        public void InconclusiveShouldBeInconclusive()
        {
            var assert = new AssertClass();
            assert.Inconclusive();

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Inconclusive);
        }

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [TestMethod]
        public void NullValuesAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, null);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Succeeded);
        }

        [TestMethod]
        public void NullIsNotComparableWithValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Inconclusive);
        }

        [TestMethod]
        public void DifferentTypesEqualityIsNotConclusive()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, "42");

            Assert.AreEqual(_args.Result.Outcome, TestOutcome.Inconclusive);
        }
    }
}
