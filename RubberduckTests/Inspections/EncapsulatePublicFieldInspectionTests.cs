using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class EncapsulatePublicFieldInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void PublicField_ReturnsResult()
        {
            const string inputCode =
@"Public fizz As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void MultiplePublicFields_ReturnMultipleResult()
        {
            const string inputCode =
@"Public fizz As Boolean
Public buzz As Integer, _
       bazz As Integer";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(3, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void PrivateField_DoesNotReturnResult()
        {
            const string inputCode =
@"Private fizz As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void PublicNonField_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByRef arg1 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void PublicField_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore EncapsulatePublicField
Public fizz As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EncapsulatePublicField_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public fizz As Boolean";

            const string expectedCode =
@"'@Ignore EncapsulatePublicField
Public fizz As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new EncapsulatePublicFieldInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "EncapsulatePublicFieldInspection";
            var inspection = new EncapsulatePublicFieldInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
