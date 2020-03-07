using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnhandledOnErrorResumeNextInspectionTest : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_GoToDeclaredBefore_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error GoTo 0
    On Error Resume Next
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_GoToDeclaredAfter_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
    On Error GoTo 0
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_MultipleStatements_ReturnsMultipleResults()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
    On Error Resume Next
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_GoToDeclaredBeforeMultipleStatements_ReturnsMultipleResults()
        {
            const string inputCode =
@"Sub Foo()
    On Error GoTo 0
    On Error Resume Next
    On Error Resume Next
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_GoToDeclaredBetweenMultipleStatements_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
    On Error GoTo 0
    On Error Resume Next
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_GoToDeclaredAfterMultipleStatements_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
    On Error Resume Next
    On Error GoTo 0
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_StatementsInDifferentProcedures_ReturnsResultsPerProcedure()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
    On Error GoTo 0
End Sub

Sub Bar()
    On Error Resume Next
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new UnhandledOnErrorResumeNextInspection(null);

            Assert.AreEqual(nameof(UnhandledOnErrorResumeNextInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UnhandledOnErrorResumeNextInspection(state);
        }
    }
}

