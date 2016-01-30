using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
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

        [TestMethod, Timeout(1000)]
        public void IsTrueSucceedsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(true);
            
            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsTrueFailsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(false);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsFalseSucceedsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(false);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsFalseFailsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(true);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreSameShouldSucceedWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = obj1;
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreSameShouldFailWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = new object();
            assert.AreSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreSameShouldSucceedWithTwoNullReferences()
        {
            var assert = new AssertClass();
            assert.AreSame(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreSameShouldFailWithActualNullReference()
        {
            var assert = new AssertClass();
            assert.AreSame(new object(), null);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreSameShouldFailWithExpectedNullReference()
        {
            var assert = new AssertClass();
            assert.AreSame(null, new object());

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotSameShouldSucceedWithDifferentReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = new object();
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotSameShouldSuccedWithOneNullReference()
        {
            var assert = new AssertClass();
            assert.AreNotSame(new object(), null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotSameShouldFailWithBothNullReferences()
        {
            var assert = new AssertClass();
            assert.AreNotSame(null, null);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotSameShouldFailWithSameReferences()
        {
            var assert = new AssertClass();
            var obj1 = new object();
            var obj2 = obj1;
            assert.AreNotSame(obj1, obj2);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreEqualShouldSucceedWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 1);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreEqualShouldFailWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 2);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotEqualShouldSucceedWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotEqualShouldFailWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 1);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void AreNotEqualShouldBeInconclusiveWithDifferentTypes()
        {
            int obj1 = 10;
            double obj2 = 10;

            var assert = new AssertClass();
            assert.AreNotEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsNothingShouldSucceedWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsNothingShouldFailWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(new object());

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsNotNothingShouldFailWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(null);

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void IsNotNothingShouldSucceedWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(new object());

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void FailShouldFail()
        {
            var assert = new AssertClass();
            assert.Fail();

            Assert.AreEqual(TestOutcome.Failed, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void InconclusiveShouldBeInconclusive()
        {
            var assert = new AssertClass();
            assert.Inconclusive();

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Result.Outcome);
        }

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [TestMethod, Timeout(1000)]
        public void NullValuesAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void NullAndEmptyStringAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, string.Empty);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void NullIsNotComparableWithValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Result.Outcome);
        }

        [TestMethod, Timeout(1000)]
        public void DifferentTypesEqualityIsNotConclusive()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, "42");

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Result.Outcome);
        }
    }
}
