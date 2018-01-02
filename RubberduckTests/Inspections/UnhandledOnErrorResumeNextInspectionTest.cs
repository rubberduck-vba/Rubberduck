using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnhandledOnErrorResumeNextInspectionTest
    {
        [Test]
        [Category("Inspections")]
        public void UnhandledOnErrorResumeNext_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    On Error Resume Next
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new UnhandledOnErrorResumeNextInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UnhandledOnErrorResumeNextInspection";
            var inspection = new UnhandledOnErrorResumeNextInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

