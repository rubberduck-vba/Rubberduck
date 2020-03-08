using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteErrorSyntaxInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteErrorSyntax_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Error 91
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteErrorSyntax_DoesNotReturnResult_ErrorInStringLiteral()
        {
            const string inputCode =
                @"Sub Foo()
    Dim bar As String
    bar = ""Error 91"" ' test
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteErrorSyntax_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Sub Foo()
    Error 91
    Error 91
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteErrorSyntax_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    '@Ignore ObsoleteErrorSyntax
    Error 91
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteErrorSyntaxInspection(null);

            Assert.AreEqual(nameof(ObsoleteErrorSyntaxInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteErrorSyntaxInspection(state);
        }
    }
}
