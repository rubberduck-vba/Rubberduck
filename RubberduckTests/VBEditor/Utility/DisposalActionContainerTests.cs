using NUnit.Framework;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.VBEditor.Utility
{
    [TestFixture()]
    public class DisposalActionContainerTests
    {
        [Test()]
        public void ValueReturnsValuePassedIn()
        {
            var testValue = 42;
            var dac = DisposalActionContainer.Create(testValue, (a) => { });
            var returnedValue = dac.Value;

            Assert.AreEqual(testValue, returnedValue);
        }

        [Test()]
        public void FirstDisposeTriggersActionPassedIn()
        {
            var useCount = 0;
            var dac = DisposalActionContainer.Create(42, (a) => useCount++);
            dac.Dispose();
            var expectedUseCount = 1;

            Assert.AreEqual(expectedUseCount, useCount);
        }

        [Test()]
        public void MultipleCallsOfDisposeTriggerTheActionPassedInOnce()
        {
            var useCount = 0;
            var dac = DisposalActionContainer.Create(42, (a) => useCount++);
            dac.Dispose();
            dac.Dispose();
            dac.Dispose();
            dac.Dispose();
            dac.Dispose();
            dac.Dispose();
            var expectedUseCount = 1;

            Assert.AreEqual(expectedUseCount, useCount);
        }

        [Test()]
        public void ArgumentForActionPassedInIsTheValuePassedIn()
        {
            var actualArgumentValue = 0;
            var dac = DisposalActionContainer.Create(42, (a) => actualArgumentValue = a);
            dac.Dispose();
            var expectedArgumentValue = 42;

            Assert.AreEqual(expectedArgumentValue, actualArgumentValue);
        }
    }
}