using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.Extensions;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ImplementInterfaceTests
    {
        [TestMethod]
        public void ImplementInterface_Procedure()
        {
            //Input
            const string inputCode1 =
@"Public Sub Foo()
End Sub";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_Procedure_WithParams()
        {
            //Input
            const string inputCode1 =
@"Public Sub Foo(ByVal a As Integer, ByRef b, c, d As Long)
End Sub";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Sub Class1_Foo(ByVal a As Integer, ByRef b As Variant, ByRef c As Variant, ByRef d As Long)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_Function()
        {
            //Input
            const string inputCode1 =
@"Public Function Foo() As Integer
End Function";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Function Class1_Foo() As Integer
    Err.Raise 5 'TODO implement interface member
End Function
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_Function_WithImplicitType()
        {
            //Input
            const string inputCode1 =
@"Public Function Foo()
End Function";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Function Class1_Foo() As Variant
    Err.Raise 5 'TODO implement interface member
End Function
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_Function_WithParam()
        {
            //Input
            const string inputCode1 =
@"Public Function Foo(a)
End Function";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Function Class1_Foo(ByRef a As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertyGet()
        {
            //Input
            const string inputCode1 =
@"Public Property Get Foo() As Integer
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Get Class1_Foo() As Integer
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertyGet_WithImplicitType()
        {
            //Input
            const string inputCode1 =
@"Public Property Get Foo()
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Get Class1_Foo() As Variant
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertyGet_WithParam()
        {
            //Input
            const string inputCode1 =
@"Public Property Get Foo(a)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Get Class1_Foo(ByRef a As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertyLet()
        {
            //Input
            const string inputCode1 =
@"Public Property Let Foo(ByRef value As Long)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Let Class1_Foo(ByRef value As Long)
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertyLet_WithParam()
        {
            //Input
            const string inputCode1 =
@"Public Property Let Foo(a)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Let Class1_Foo(ByRef a As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertySet()
        {
            //Input
            const string inputCode1 =
@"Public Property Set Foo(ByRef value As Variant)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Set Class1_Foo(ByRef value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertySet_WithParam()
        {
            //Input
            const string inputCode1 =
@"Public Property Set Foo(a)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Property Set Class1_Foo(ByRef a As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplementInterface_PropertySet_AllTypes()
        {
            //Input
            const string inputCode1 =
@"Public Sub Foo()
End Sub

Public Function Bar(ByVal a As Integer) As Boolean
End Function

Public Property Get Buz(ByVal a As Boolean) As Integer
End Property

Public Property Let Buz(ByVal a As Boolean, ByRef value As Integer)
End Property";

            const string inputCode2 =
@"Implements Class1";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Implements Class1

Private Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function Class1_Bar(ByVal a As Integer) As Boolean
    Err.Raise 5 'TODO implement interface member
End Function

Private Property Get Class1_Buz(ByVal a As Boolean) As Integer
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let Class1_Buz(ByVal a As Boolean, ByRef value As Integer)
    Err.Raise 5 'TODO implement interface member
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                 .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                 .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(1);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module = project.Object.VBComponents.Item(1).CodeModule;

            //Act
            var refactoring = new ImplementInterfaceRefactoring(vbe.Object, parser.State, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }
    }
}