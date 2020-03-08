using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MultipleDeclarationsInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_ReturnsResult_Variables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_ReturnsResult_Constants()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const var1 As Integer = 9, var2 As String = ""test""
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_ReturnsResult_StaticVariables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Static var1 As Integer, var2 As String
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_ReturnsResult_MultipleDeclarations()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
    Dim var3 As Boolean, var4 As Date
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_ReturnsResult_SomeDeclarationsSeparate()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
    Dim var3 As Boolean
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void MultipleDeclarations_Ignore_DoesNotReturnResult_Variables()
        {
            const string inputCode =
                @"Public Sub Foo()
    '@Ignore MultipleDeclarations
    Dim var1 As Integer, var2 As String
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new MultipleDeclarationsInspection(null);

            Assert.AreEqual(nameof(MultipleDeclarationsInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MultipleDeclarationsInspection(state);
        }
    }
}
