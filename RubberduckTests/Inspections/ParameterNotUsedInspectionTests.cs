using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ParameterNotUsedInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ParameterUsed_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
    arg1 = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_SomeParamsUsed()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer, ByVal arg2 as String)
    arg1 = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults().ToList();

                Assert.AreEqual(1, inspectionResults.Count);

            }
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ParameterNotUsed
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new ParameterNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ParameterNotUsedInspection";
            var inspection = new ParameterNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
