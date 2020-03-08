using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitPublicMemberInspectionTests : InspectionTestsBase
    {
        [TestCase("Sub Foo()\r\nEnd Sub", 1)]
        [TestCase("Function Foo() As Boolean\r\nEnd Function", 1)]
        [TestCase("Sub Foo()\r\nEnd Sub\r\n\r\nSub Goo()\r\nEnd Sub", 2)]
        [TestCase("Private Sub Foo()\r\nEnd Sub\r\n\r\nSub Goo()\r\nEnd Sub", 1)]
        [TestCase("Private Sub Foo()\r\nEnd Sub", 0)]
        [TestCase("'@Ignore ImplicitPublicMember\r\nSub Foo()\r\nEnd Sub", 0)]
        [Category("Inspections")]
        public void ImplicitPublicMember_Various(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitPublicMemberInspection(null);

            Assert.AreEqual(nameof(ImplicitPublicMemberInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitPublicMemberInspection(state);
        }
    }
}
