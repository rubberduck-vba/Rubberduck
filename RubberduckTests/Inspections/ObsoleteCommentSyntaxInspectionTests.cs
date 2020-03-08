using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteCommentSyntaxInspectionTests : InspectionTestsBase
    {
        [TestCase("Rem test", 1)]
        [TestCase("' test", 0)] //QuoteComment
        [TestCase("Rem test1\r\nRem test2", 2)]
        [TestCase("Rem test1\r\n' test2", 1)] //SomeObsoleteCommentSyntax
        [TestCase("'@Ignore ObsoleteCommentSyntax\r\nRem test", 0)]
        [TestCase("Rem test", 1)]
        [Category("Inspections")]
        public void ObsoleteCommentSyntax_VariousScenarios(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_OtherParseInspectionFires()
        {
            const string inputCode = @"
Sub foo()
    Dim i As String
    i = """"
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_RemInStringLiteral()
        {
            const string inputCode =
                @"Sub Foo()
    Dim bar As String
    bar = ""iejo rem oernp"" ' test
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCommentSyntaxInspection";
            var inspection = new ObsoleteCommentSyntaxInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteCommentSyntaxInspection(state);
        }
    }
}
