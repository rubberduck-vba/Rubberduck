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
        [TestCase("Enum Foo\r\n    Member\r\nEnd Enum", 1)]
        [TestCase("Enum Foo\r\n    Member\r\nEnd Enum\r\n\r\nEnum Bar\r\n    Member\r\nEnd Enum", 2)]
        [TestCase("Private Enum Foo\r\n    Member\r\nEnd Enum\r\n\r\nEnum Bar\r\n    Member\r\nEnd Enum", 1)]
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

        [TestCase(null, 1)]
        [TestCase(true, 0)]
        [TestCase(false, 0)]
        [Category(nameof(ImplicitPublicMemberInspectionTests))]
        [Category("Inspections")]
        public void User_defined_type_correctly_flagged_as_implicitly_public(bool? modifier, int expectedCount)
        {
            string accessModifier = null;
            if (!modifier.HasValue)
            {
                accessModifier = string.Empty;
            }
            else
            {
                accessModifier = modifier.Value ? "Public " : "Private ";
            }

            var code = $@"Option Explicit
{accessModifier}Type FooType
    Name As String
End Type
";

            var inspectionResults = InspectionResultsForStandardModule(code);

            var actual = inspectionResults.Count();

            Assert.AreEqual(expectedCount, actual);
        }
    }
}
