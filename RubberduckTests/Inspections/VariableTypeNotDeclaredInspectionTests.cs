using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class VariableTypeNotDeclaredInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleParams()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Parameters()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2 As String)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Variables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2 As Date
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As Integer
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void VariableTypeNotDeclared_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore VariableTypeNotDeclared
Sub Foo(arg1)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new VariableTypeNotDeclaredInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "VariableTypeNotDeclaredInspection";
            var inspection = new VariableTypeNotDeclaredInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
