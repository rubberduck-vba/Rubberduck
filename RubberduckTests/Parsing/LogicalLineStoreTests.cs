using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Parsing
{
    [TestFixture]
    public class LogicalLineStoreTests
    {
        [Test]
        [Category("Resolver")]
        [TestCase(-42)]
        [TestCase(0)]

        public void LogicalLine_ReturnsNullForPhysicalLinesSmallerThanOne(int physicalLine)
        {
            var logicalLineEnds = new List<int>{ 22, 4, 94, 28323, 3, 5, 17};
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualLogicalLineNumber = logicalLineStore.LogicalLineNumber(physicalLine);

            Assert.IsNull(actualLogicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(28324)]
        [TestCase(283324)]

        public void LogicalLine_ReturnsNullForPhysicalLinesLargerThenTheMaxPhysicalEndLine(int physicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualLogicalLineNumber = logicalLineStore.LogicalLineNumber(physicalLine);

            Assert.IsNull(actualLogicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(1, 1)]
        [TestCase(2, 1)]
        [TestCase(5, 3)]
        [TestCase(77, 6)]
        [TestCase(100, 7)]
        [TestCase(10000, 7)]
        [TestCase(17, 4)]
        [TestCase(18, 5)]

        public void LogicalLine_ReturnsExpectedLogicalLineForPhysicalLinesBetweenOneAndTheMaxPhysicalEndLine(int physicalLine, int expectedLogicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualLogicalLineNumber = logicalLineStore.LogicalLineNumber(physicalLine);

            Assert.AreEqual(expectedLogicalLine, actualLogicalLineNumber);
        }

        [Test]
        [Category("Resolver")]

        public void NumberOfLogicalLines_ReturnsNumberOfLineEnds()
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var expectedNumberOfLines = logicalLineEnds.Count;
            var actualNumberOfLines = logicalLineStore.NumberOfLogicalLines();

            Assert.AreEqual(expectedNumberOfLines, actualNumberOfLines);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(-42)]
        [TestCase(0)]

        public void PhysicalStartLine_ReturnsNullForLogicalLinesSmallerThanOne(int logicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalStartLineNumber(logicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(8)]
        [TestCase(9)]
        [TestCase(33)]

        public void PhysicalStartLine_ReturnsNullForLogicalLinesLargerThenTheNumberOfLogicalLines(int logicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalStartLineNumber(logicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(1, 1)]
        [TestCase(2, 4)]
        [TestCase(3, 5)]
        [TestCase(7, 95)]

        public void PhysicalStartLine_ReturnsExpectedPhysicalLineForExistingLogicalLines(int logicalLine, int expectedPhysicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalStartLineNumber(logicalLine);

            Assert.AreEqual(expectedPhysicalLine, actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(-42)]
        [TestCase(0)]

        public void PhysicalEndLine_ReturnsNullForLogicalLinesSmallerThanOne(int logicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalEndLineNumber(logicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(8)]
        [TestCase(9)]
        [TestCase(33)]

        public void PhysicalEndLine_ReturnsNullForLogicalLinesLargerThenTheNumberOfLogicalLines(int logicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalEndLineNumber(logicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(1, 3)]
        [TestCase(2, 4)]
        [TestCase(3, 5)]
        [TestCase(6, 94)]
        [TestCase(7, 28323)]

        public void PhysicalEndLine_ReturnsExpectedPhysicalLineForExistingLogicalLines(int logicalLine, int expectedPhysicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.PhysicalEndLineNumber(logicalLine);

            Assert.AreEqual(expectedPhysicalLine, actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(-42)]
        [TestCase(0)]

        public void StartOfContainingLogicalLine_ReturnsNullForPhysicalLinesSmallerThanOne(int physicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.StartOfContainingLogicalLine(physicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(28324)]
        [TestCase(283324)]

        public void StartOfContainingLogicalLine_ReturnsNullForPhysicalLinesLargerThenTheMaxPhysicalEndLine(int physicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.StartOfContainingLogicalLine(physicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(1, 1)]
        [TestCase(2, 1)]
        [TestCase(5, 5)]
        [TestCase(77, 23)]
        [TestCase(100, 95)]
        [TestCase(10000, 95)]
        [TestCase(17, 6)]
        [TestCase(18, 18)]

        public void StartOfContainingLogicalLine_ReturnsExpectedPhysicalLineForPhysicalLinesBetweenOneAndTheMaxPhysicalEndLine(int physicalLine, int expectedPhysicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.StartOfContainingLogicalLine(physicalLine);

            Assert.AreEqual(expectedPhysicalLine, actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(-42)]
        [TestCase(0)]

        public void EndOfContainingLogicalLine_ReturnsNullForPhysicalLinesSmallerThanOne(int physicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.EndOfContainingLogicalLine(physicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(28324)]
        [TestCase(283324)]

        public void EndOfContainingLogicalLine_ReturnsNullForPhysicalLinesLargerThenTheMaxPhysicalEndLine(int physicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.EndOfContainingLogicalLine(physicalLine);

            Assert.IsNull(actualPhysicalLineNumber);
        }

        [Test]
        [Category("Resolver")]
        [TestCase(1, 3)]
        [TestCase(2, 3)]
        [TestCase(5, 5)]
        [TestCase(77, 94)]
        [TestCase(100, 28323)]
        [TestCase(10000, 28323)]
        [TestCase(17, 17)]
        [TestCase(18, 22)]

        public void EndOfContainingLogicalLine_ReturnsExpectedPhysicalLineForPhysicalLinesBetweenOneAndTheMaxPhysicalEndLine(int physicalLine, int expectedPhysicalLine)
        {
            var logicalLineEnds = new List<int> { 22, 4, 94, 28323, 3, 5, 17 };
            var logicalLineStore = new LogicalLineStore(logicalLineEnds);

            var actualPhysicalLineNumber = logicalLineStore.EndOfContainingLogicalLine(physicalLine);

            Assert.AreEqual(expectedPhysicalLine, actualPhysicalLineNumber);
        }
    }
}