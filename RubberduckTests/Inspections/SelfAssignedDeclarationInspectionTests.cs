using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SelfAssignedDeclarationInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void SelfAssignedDeclaration_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Dim b As New Collection
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void SelfAssignedDeclaration_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Dim b As Collection
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void SelfAssignedDeclaration_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    '@Ignore SelfAssignedDeclaration
    Dim b As New Collection
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new SelfAssignedDeclarationInspection(null);

            Assert.AreEqual(nameof(SelfAssignedDeclarationInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new SelfAssignedDeclarationInspection(state);
        }
    }
}
