using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ParameterCanBeByValInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedByNotSpecified()
        {
            const string inputCode =
@"Sub Foo(arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedByRef_Unassigned()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_Multiple()
        {
            const string inputCode =
@"Sub Foo(arg1 As String, arg2 As Date)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedByValExplicitly()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedByRefAndAssigned()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_BuiltInEventParam()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_SomeParams()
        {
            const string inputCode =
@"Sub Foo(arg1 As String, ByVal arg2 As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenArrayParameter_ReturnsNoResult()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1() As Variant)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);

            var results = inspection.GetInspectionResults().ToList();

            Assert.AreEqual(0, results.Count);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByRefProc_NoAssignment()
        {
            const string inputCode =
@"Sub DoSomething(foo As Integer)
    DoSomethingElse foo
End Sub

Sub DoSomethingElse(ByVal bar As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_DoesNotReturnResult_PassedToByRefProc_WithAssignment()
        {
            const string inputCode =
@"Sub DoSomething(foo As Integer)
    DoSomethingElse foo
End Sub

Sub DoSomethingElse(ByRef bar As Integer)
    bar = 42
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_ReturnsResult_PassedToByValProc_WithAssignment()
        {
            const string inputCode =
@"Sub DoSomething(foo As Integer)
    DoSomethingElse foo
End Sub

Sub DoSomethingElse(ByVal bar As Integer)
    bar = 42
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ParameterCanBeByVal
Sub Foo(arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParam()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleByValParam()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_SingleParamUsedByRef()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
    a = 42
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_MultipleParams_OneCanBeByVal()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string inputCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();
            
            Assert.AreEqual("a", inspectionResults.Single().Target.IdentifierName);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParam()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleByValParam()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_EventMember_SingleParamUsedByRef()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_EventMember_MultipleParams_OneCanBeByVal()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByRef arg1 As Integer, ByRef arg2 As Integer)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer, ByRef arg2 As Integer)
    arg1 = 42
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual("arg2", inspectionResults.Single().Target.IdentifierName);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWorks_SubNameStartsWithParamName()
        {
            const string inputCode =
@"Sub foo(f)
End Sub";

            const string expectedCode =
@"Sub foo(ByVal f)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByUnspecified()
        {
            const string inputCode =
@"Sub Foo(arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByRef()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByUnspecified_MultilineParameter()
        {
            const string inputCode =
@"Sub Foo( _
arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo( _
ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByRef_MultilineParameter()
        {
            const string inputCode =
@"Sub Foo(ByRef _
arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal _
arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_InterfaceMember_MultipleParams_OneCanBeByVal_QuickFixWorks()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string inputCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";

            //Expected
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByRef b As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string expectedCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByRef b As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();

            var module1 = project.Object.VBComponents["IClass1"].CodeModule;
            var module2 = project.Object.VBComponents["Class1"].CodeModule;
            var module3 = project.Object.VBComponents["Class2"].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.Single().QuickFixes.Single(s => s is PassParameterByValueQuickFix).Fix();

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
            Assert.AreEqual(expectedCode3, module3.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_EventMember_MultipleParams_OneCanBeByVal_QuickFixWorks()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByRef a As Integer, ByRef b As Integer)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByRef b As Integer)
    a = 42
End Sub";
            const string inputCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByRef b As Integer)
End Sub";

            //Expected
            const string expectedCode1 =
@"Public Event Foo(ByRef a As Integer, ByVal b As Integer)";
            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByVal b As Integer)
    a = 42
End Sub";
            const string expectedCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByVal b As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode3)
                .Build();

            var module1 = project.Object.VBComponents["Class1"].CodeModule;
            var module2 = project.Object.VBComponents["Class2"].CodeModule;
            var module3 = project.Object.VBComponents["Class3"].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.Single().QuickFixes.Single(s => s is PassParameterByValueQuickFix).Fix();

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
            Assert.AreEqual(expectedCode3, module3.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(ByRef _
arg1 As String)
End Sub";

            const string expectedCode =
@"'@Ignore ParameterCanBeByVal
Sub Foo(ByRef _
arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWithOptionalWorks()
        {
            const string inputCode =
@"Sub Test(Optional foo As String = ""bar"")
    Debug.Print foo
End Sub";

            const string expectedCode =
@"Sub Test(Optional ByVal foo As String = ""bar"")
    Debug.Print foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.Single(s => s is PassParameterByValueQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWithOptionalByRefWorks()
        {
            const string inputCode =
@"Sub Test(Optional ByRef foo As String = ""bar"")
    Debug.Print foo
End Sub";

            const string expectedCode =
@"Sub Test(Optional ByVal foo As String = ""bar"")
    Debug.Print foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.Single(s => s is PassParameterByValueQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [TestMethod]
        [TestCategory("Inspections")]
        public void ParameterCanBeByVal_QuickFixWithOptional_LineContinuationsWorks()
        {
            const string inputCode =
@"Sub foo(Optional _
  ByRef _
  foo _
  As _
  Byte _
  )
  Debug.Print foo
End Sub";

            const string expectedCode =
@"Sub foo(Optional _
  ByVal _
  foo _
  As _
  Byte _
  )
  Debug.Print foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects[0];
            var module = project.VBComponents[0].CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ParameterCanBeByValInspection(parser.State);
            inspection.GetInspectionResults().First().QuickFixes.Single(s => s is PassParameterByValueQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ParameterCanBeByValInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ParameterCanBeByValInspection";
            var inspection = new ParameterCanBeByValInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
