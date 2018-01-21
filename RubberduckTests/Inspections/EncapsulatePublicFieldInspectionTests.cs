using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EncapsulatePublicFieldInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void PublicField_ReturnsResult()
        {
            const string inputCode =
                @"Public fizz As Boolean";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MultiplePublicFields_ReturnMultipleResult()
        {
            const string inputCode =
                @"Public fizz As Boolean
Public buzz As Integer, _
       bazz As Integer";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(3, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void PrivateField_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private fizz As Boolean";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void PublicNonField_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void PublicField_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore EncapsulatePublicField
Public fizz As Boolean";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new EncapsulatePublicFieldInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "EncapsulatePublicFieldInspection";
            var inspection = new EncapsulatePublicFieldInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
