using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class NonReturningFunctionInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningPropertyGet_ReturnsResult()
        {
            const string inputCode =
                @"Property Get Foo() As Boolean
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_GivenParenthesizedByRefAssignment()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    ByRefAssign (Foo)
End Function

Public Sub ByRefAssign(ByRef a As Boolean)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenByRefAssignment_WithMemberAccess()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    TestModule1.ByRefAssign False, Foo
End Function

Public Sub ByRefAssign(ByVal v As Boolean, ByRef a As Boolean)
    a = v
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_ReturnsResult_GivenUnassignedByRefAssignment_WithMemberAccess()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    TestModule1.ByRefAssign False, Foo
End Function

Public Sub ByRefAssign(ByVal v As Boolean, ByRef a As Boolean)
    'nope, not assigned
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenClassMemberCall()
        {
            const string code = @"
Public Function Foo() As Boolean
    With New Class1
        .ByRefAssign Foo
    End With
End Function
";
            const string classCode = @"
Public Sub ByRefAssign(ByRef b As Boolean)
End Sub
";
            var builder = new MockVbeBuilder();
            builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule1", ComponentType.StandardModule, code)
                .AddComponent("Class1", ComponentType.ClassModule, classCode);
            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonReturningFunction_NoResult_GivenByRefAssignment_WithNamedArgument()
        {
            const string inputCode = @"
Public Function Foo() As Boolean
    ByRefAssign b:=Foo
End Function

Public Sub ByRefAssign(Optional ByVal a As Long, Optional ByRef b As Boolean)
    b = False
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "NonReturningFunctionInspection";
            var inspection = new NonReturningFunctionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
