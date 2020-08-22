using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyStringLiteralInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyStringLiteral_ReturnsResult_PassToProcedure()
        {
            const string inputCode =
                @"Public Sub Bar()
    Foo """"
End Sub

Public Sub Foo(ByRef arg1 As String)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyStringLiteral_ReturnsResult_Assignment()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
    arg1 = """"
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NotEmptyStringLiteral_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyStringLiteral_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
    '@Ignore EmptyStringLiteral
    arg1 = """"
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new EmptyStringLiteralInspection(null);

            Assert.AreEqual(nameof(EmptyStringLiteralInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyStringLiteralInspection(state);
        }
    }
}
