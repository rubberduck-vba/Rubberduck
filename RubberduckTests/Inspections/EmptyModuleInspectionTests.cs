using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyModuleInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ModuleWithContentNotRepresentingFunctionality_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Option Compare Binary
Option Explicit
Option Private Module

'Nothing to see here. 

DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V

'Here, neither! _
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ClassWithContentNotRepresentingFunctionality_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Option Compare Binary
Option Explicit

'Nothing to see here. 

DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V

'Here, neither! _
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithFunction_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Function Foo() As String
End Function
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithProcedure_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithPropertyGet_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Property Get Foo()
End Property
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithPropertySet_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Property Set Foo(rhs As Variant)
End Property
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithPropertyLet_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Property Let Foo(rhs As Variant)
End Property
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.UserForm, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithEnum_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Enum Foo
Bar
End Enum
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithPrivateUDT_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Type Foo
Bar As String
End Type
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithPublicUDT_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Type Foo
Bar As String
End Type
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithInstanceVariable_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private foo As String";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithConstant_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Const foo As String = """"";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleWithEvent_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Event Foo(bar As Variant)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyDocumentModule_DoesNotReturnResult()
        {
            const string inputCode = "";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.Document, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyUserForm_DoesNotReturnResult()
        {
            const string inputCode = "";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.Document, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void EmptyModule_Ignored_DoesNotReturnResult()
        {
            const string inputCode = "'@Ignore EmptyModule";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyModuleInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }


        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new EmptyModuleInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "EmptyModuleInspection";
            var inspection = new EmptyModuleInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
