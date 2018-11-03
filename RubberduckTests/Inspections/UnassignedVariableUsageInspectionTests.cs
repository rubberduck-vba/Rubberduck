using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnassignedVariableUsageInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_ReturnsResult()
        {
            const string inputCode = @"
Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        // this test will eventually be removed once we can fire the inspection on a specific reference
        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_ReturnsSingleResult_MultipleReferences()
        {
            const string inputCode =
                @"Sub tester()
    Dim myarr() As Variant
    Dim i As Long

    ReDim myarr(1 To 10)

    For i = 1 To 10
        DoSomething myarr(i)
    Next

End Sub

Sub DoSomething(ByVal foo As Variant)
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    b = True
    bb = b
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()    
    Dim b As Boolean
    Dim bb As Boolean

'@Ignore UnassignedVariableUsage
    bb = b
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_Ignored_DoesNotReturnResultMultipleIgnores()
        {
            const string inputCode =
                @"Sub Foo()    
    Dim b As Boolean
    Dim bb As Boolean

'@Ignore UnassignedVariableUsage, VariableNotAssigned
    bb = b
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultIfNoReferences()
        {
            const string inputCode =
                @"Sub DoSomething()
    Dim foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultForLenFunction()
        {
            const string inputCode =
                @"Sub DoSomething()
    Dim foo As LongPtr
    Debug.Print Len(foo)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultForLenBFunction()
        {
            const string inputCode =
                @"Sub DoSomething()
    Dim foo As LongPtr
    Debug.Print LenB(foo)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UnassignedVariableUsageInspection";
            var inspection = new UnassignedVariableUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
