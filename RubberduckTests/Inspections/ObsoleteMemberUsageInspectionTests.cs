using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteMemberUsageInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsed_ReturnsResult()
        {
            const string inputCode = @"
'@Obsolete
Public Sub Foo()
End Sub

Public Sub Bar()
    Foo
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsedTwice_ReturnsTwoResults()
        {
            const string inputCode = @"
'@Obsolete
Public Sub Foo()
End Sub

Public Sub Bar()
    Foo
    Foo
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteMemberUsedOnNonMemberDeclaration_DoesNotReturnResult()
        {
            const string inputCode = @"
'@Obsolete
Public s As String";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteMemberUsageInspection(state);
        }
    }
}
