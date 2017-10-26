using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class StopKeywordInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void StopKeyword_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Stop
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new StopKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NoStopKeyword_NoResult()
        {
            var inputCode =
                @"Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new StopKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void StopKeyword_Ignored_DoesNotReturnResult()
        {
            var inputCode =
                @"Sub Foo()
'@Ignore StopKeyword
    Stop
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void StopKeywords_Ignored_ReturnsCorrectResults()
        {
            var inputCode =
                @"Sub Foo()
    Dim d As Integer
    d = 0
    Stop

    d = 1

    '@Ignore StopKeyword
    Stop
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new StopKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
                Assert.AreEqual(4, inspectionResults.First().QualifiedSelection.Selection.StartLine);
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new StopKeywordInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(StopKeywordInspection);
            var inspection = new StopKeywordInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
