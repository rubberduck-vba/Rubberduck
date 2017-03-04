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
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class WriteOnlyPropertyInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_Let()
        {
            const string inputCode = 
@"Property Let Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_Set()
        {
            const string inputCode =
@"Property Set Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_LetAndSet()
        {
            const string inputCode =
@"Property Let Foo(value)
End Property

Property Set Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_DoesNotReturnsResult_Get()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_DoesNotReturnsResult_GetAndLetAndSet()
        {
            const string inputCode =
@"Property Get Foo()
End Property

Property Let Foo(value)
End Property

Property Set Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore WriteOnlyProperty
Property Let Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_ImplicitTypesAndAccessibility()
        {
            const string inputCode =
@"Property Let Foo(value)
End Property";

            const string expectedCode =
@"Public Property Get Foo() As Variant
End Property
Property Let Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is WriteOnlyPropertyQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_ExlicitTypesAndAccessibility()
        {
            const string inputCode =
@"Public Property Let Foo(ByVal value As Integer)
End Property";

            const string expectedCode =
@"Public Property Get Foo() As Integer
End Property
Public Property Let Foo(ByVal value As Integer)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is WriteOnlyPropertyQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_MultipleParams()
        {
            const string inputCode =
@"Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            const string expectedCode =
@"Public Property Get Foo(ByRef value1 As Variant, ByVal value2 As Integer, ByRef value3 As Long, ByRef value4 As Date, ByVal value5 As Variant) As String
End Property
Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is WriteOnlyPropertyQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void WriteOnlyProperty_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Property Let Foo(value)
End Property";

            const string expectedCode =
@"'@Ignore WriteOnlyProperty
Property Let Foo(value)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var module = project.Object.VBComponents[0].CodeModule;
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new WriteOnlyPropertyInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, module.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new WriteOnlyPropertyInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "WriteOnlyPropertyInspection";
            var inspection = new WriteOnlyPropertyInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
