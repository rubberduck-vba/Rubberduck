using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class NonReturningFunctionInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningPropertyGet_ReturnsResult()
        {
            const string inputCode =
                @"Property Get Foo() As Boolean
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleFunctions()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function

Function Goo() As String
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_DoesNotReturnResult_Let()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Foo = True
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_DoesNotReturnResult_Set()
        {
            const string inputCode =
                @"Function Foo() As Collection
    Set Foo = new Collection
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore NonReturningFunction
Function Foo() As Boolean
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleSubs_SomeReturning()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Foo = True
End Function

Function Goo() As String
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
                @"Function Foo() As Boolean
End Function";
            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Boolean
End Function";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new NonReturningFunctionInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "NonReturningFunctionInspection";
            var inspection = new NonReturningFunctionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
