using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MultilineParameterInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void MultilineParameter_ReturnsResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultilineParameter_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal Var1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultilineParameter_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer, _
    ByVal _
    Var2 _
    As _
    Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultilineParameter_ReturnsResults_SomeParams()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer, ByVal Var2 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultilineParameter_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore MultilineParameter
Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new MultilineParameterInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "MultilineParameterInspection";
            var inspection = new MultilineParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
