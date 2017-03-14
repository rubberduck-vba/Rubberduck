using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    //NOTE: The tests for reference equity are commented out pending some way of figuring out how to test the correct behavior.
    //These methods have to check to see if the parameters are COM objects (see https://github.com/rubberduck-vba/Rubberduck/issues/2848)
    //to make the result match the VBA interpretations of reference and value types.  Similarly, the SequenceEqual and NotSequenceEqual
    //methods remain untested because they make several of the same Type tests that are AFAIK impossible to mock.

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

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsTrueSucceedsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(true);
            
            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsTrueFailsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsTrue(false);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsFalseSucceedsWithFalseExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(false);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsFalseFailsWithTrueExpression()
        {
            var assert = new AssertClass();
            assert.IsFalse(true);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreSameShouldSucceedWithSameReferences()
        //{
        //    var assert = new AssertClass();
        //    var obj1 = new object();
        //    var obj2 = obj1;
        //    assert.AreSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreSameShouldFailWithDifferentReferences()
        //{
        //    var assert = new AssertClass();
        //    var obj1 = new object();
        //    var obj2 = new object();
        //    assert.AreSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreSameShouldSucceedWithTwoNullReferences()
        {
            var assert = new AssertClass();
            assert.AreSame(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreSameShouldFailWithActualNullReference()
        //{
        //    var assert = new AssertClass();
        //    assert.AreSame(new object(), null);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreSameShouldFailWithExpectedNullReference()
        //{
        //    var assert = new AssertClass();
        //    assert.AreSame(null, new object());

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreNotSameShouldSucceedWithDifferentReferences()
        //{
        //    var assert = new AssertClass();
        //    var obj1 = new object();
        //    var obj2 = new object();
        //    assert.AreNotSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreNotSameShouldSuccedWithOneNullReference()
        //{
        //    var assert = new AssertClass();
        //    assert.AreNotSame(new object(), null);

        //    Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        //}

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreNotSameShouldFailWithBothNullReferences()
        {
            var assert = new AssertClass();
            assert.AreNotSame(null, null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        //[TestCategory("Unit Testing")]
        //[TestMethod]
        //public void AreNotSameShouldFailWithSameReferences()
        //{
        //    var assert = new AssertClass();
        //    var obj1 = new object();
        //    var obj2 = obj1;
        //    assert.AreNotSame(obj1, obj2);

        //    Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        //}

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreEqualShouldSucceedWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 1);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreEqualShouldFailWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(1, 2);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreNotEqualShouldSucceedWithDifferentValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 2);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreNotEqualShouldFailWithSameValues()
        {
            var assert = new AssertClass();
            assert.AreNotEqual(1, 1);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void AreNotEqualShouldBeInconclusiveWithDifferentTypes()
        {
            int obj1 = 10;
            double obj2 = 10;

            var assert = new AssertClass();
            assert.AreNotEqual(obj1, obj2);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsNothingShouldSucceedWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsNothingShouldFailWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNothing(new object());

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsNotNothingShouldFailWithNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(null);

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void IsNotNothingShouldSucceedWithNonNullValue()
        {
            var assert = new AssertClass();
            assert.IsNotNothing(new object());

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void FailShouldFail()
        {
            var assert = new AssertClass();
            assert.Fail();

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void InconclusiveShouldBeInconclusive()
        {
            var assert = new AssertClass();
            assert.Inconclusive();

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        private void AssertHandler_OnAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _args = e;
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void NullValuesAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, null);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void NullAndEmptyStringAreEqual()
        {
            var assert = new AssertClass();
            assert.AreEqual(null, string.Empty);

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void NullIsNotComparableWithValues()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, null);

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void DifferentTypesEqualityIsNotConclusive()
        {
            var assert = new AssertClass();
            assert.AreEqual(42, "42");

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void OnAssertSucceeded_ReturnsResultSuccess()
        {
            AssertHandler.OnAssertSucceeded();

            Assert.AreEqual(TestOutcome.Succeeded, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void OnAssertFailed_ReturnsResultFailed()
        {
            AssertHandler.OnAssertFailed("MyMethod", "I Failed");

            Assert.AreEqual(TestOutcome.Failed, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void OnAssertInconclusive_ReturnsResultInconclusive()
        {
            AssertHandler.OnAssertInconclusive("Inconclusive");

            Assert.AreEqual(TestOutcome.Inconclusive, _args.Outcome);
        }

        [TestCategory("Unit Testing")]
        [TestMethod]
        public void OnAssertIgnored_ReturnsResultIgnored()
        {
            AssertHandler.OnAssertIgnored();

            Assert.AreEqual(TestOutcome.Ignored, _args.Outcome);
        }
    }
}
