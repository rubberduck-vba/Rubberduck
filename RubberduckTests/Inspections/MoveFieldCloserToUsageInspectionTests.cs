using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MoveFieldCloseToUsageInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_ReturnsResult()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_MultipleReferenceInDifferentScope()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    Let bar = ""test""
End Sub
Public Sub For2()
    Let bar = ""test""
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnResult_Variable()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim bar As String
    bar = ""test""
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_NoReferences()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertyGet()
        {
            const string inputCode =
                @"Private bar As String
Public Property Get Foo() As String
    Foo = bar
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertyLet()
        {
            const string inputCode =
                @"Private bar As String
Public Property Get Foo() As String
    Foo = ""test""
End Property
Public Property Let Foo(ByVal value As String)
    bar = value
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_DoesNotReturnsResult_ReferenceInPropertySet()
        {
            const string inputCode =
                @"Private bar As Variant
Public Property Get Foo() As Variant
    Foo = ""test""
End Property
Public Property Set Foo(ByVal value As Variant)
    bar = value
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MoveFieldCloserToUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore MoveFieldCloserToUsage
Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new MoveFieldCloserToUsageInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "MoveFieldCloserToUsageInspection";
            var inspection = new MoveFieldCloserToUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
