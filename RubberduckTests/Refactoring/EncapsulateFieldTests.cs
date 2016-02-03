using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class EncapsulateFieldTests
    {
        [TestMethod]
        public void EncapsulatePublicField_WithLetter()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_WithSetter()
        {
            //Input
            const string inputCode =
@"Public fizz As Variant";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Variant

Public Property Get Name() As Variant
    Name = fizz
End Property

Public Property Set Name(ByVal value As Variant)
    fizz = value
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = false,
                ImplementSetSetterType = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_WithOnlyGetter()
        {
            //Input
            const string inputCode =
@"Public fizz As Variant";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Variant

Public Property Get Name() As Variant
    Name = fizz
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = false,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_OtherMethodsInClass()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer

Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property

Sub Foo()
End Sub

Function Bar() As Integer
    Bar = 0
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_OtherPropertiesInClass()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer

Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property

Property Get Foo() As Variant
    Foo = True
End Property

Property Let Foo(ByVal vall As Variant)
End Property

Property Set Foo(ByVal vall As Variant)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_OtherFieldsInClass()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Public buzz As Boolean";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Public buzz As Boolean
Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveFirst()
        {
            //Input
            const string inputCode =
@"Public fizz, _
         buzz As Boolean, _
         bazz As Date";
            var selection = new Selection(1, 12, 1, 12);

            //Expectation
            const string expectedCode =
@"Public          buzz As Boolean,         bazz As Date
Private fizz As Variant

Public Property Get Name() As Variant
    Name = fizz
End Property

Public Property Let Name(ByVal value As Variant)
    fizz = value
End Property

Public Property Set Name(ByVal value As Variant)
    fizz = value
End Property
";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveSecond()
        {
            //Input
            const string inputCode =
@"Public fizz, _
         buzz As Boolean, _
         bazz As Date";
            var selection = new Selection(2, 12, 2, 12);

            //Expectation
            const string expectedCode =
@"Public fizz,                  bazz As Date
Private buzz As Boolean

Public Property Get Name() As Boolean
    Name = buzz
End Property

Public Property Let Name(ByVal value As Boolean)
    buzz = value
End Property
";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveLast()
        {
            //Input
            const string inputCode =
@"Public fizz, _
         buzz As Boolean, _
         bazz As Date";
            var selection = new Selection(3, 12, 3, 12);

            //Expectation
            const string expectedCode =
@"Public fizz,         buzz As Boolean         
Private bazz As Date

Public Property Get Name() As Date
    Name = bazz
End Property

Public Property Let Name(ByVal value As Date)
    bazz = value
End Property
";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePrivateField()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldHasReferences()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer

Sub Foo()
    fizz = 0
    Bar fizz
End Sub

Sub Bar(ByVal name As Integer)
End Sub";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property

Sub Foo()
    Name = 0
    Bar Name
End Sub

Sub Bar(ByVal name As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldHasReferencesInMultipleClasses()
        {
            //Input
            const string inputCode1 =
@"Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            const string inputCode2 =
@"Sub Foo()
    Dim c As Class1
    c.fizz = 0
    Bar c.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode1 =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property

Sub Foo()
    Name = 1
End Sub";

            const string expectedCode2 =
@"Sub Foo()
    Dim c As Class1
    c.Name = 0
    Bar c.Name
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void EncapsulatePublicField_PassInTarget()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";
            var selection = new Selection(1, 1, 1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Integer

Public Property Get Name() As Integer
    Name = fizz
End Property

Public Property Let Name(ByVal value As Integer)
    fizz = value
End Property
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(parser.State, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(parser.State.AllUserDeclarations.FindVariable(qualifiedSelection));

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, null);

            //act
            var refactoring = new EncapsulateFieldRefactoring(factory, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParams_ModelIsNull()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 1, 1, 1);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //SetupFactory
            var factory = SetupFactory(null);

            //Act
            var refactoring = new EncapsulateFieldRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void Factory_NullSelectionReturnsNullPresenter()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, null);

            Assert.AreEqual(null, factory.Create());
        }

        [TestMethod]
        public void Presenter_ParameterlessTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var codePane = project.Object.VBComponents.Item(0).CodeModule.CodePane;
            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, null);
            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_NullTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = Selection.Home;

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var codePane = project.Object.VBComponents.Item(0).CodeModule.CodePane;
            var ext = codePaneFactory.Create(codePane);
            ext.Selection = selection;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithParameterNameChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);
            view.SetupProperty(v => v.ParameterName, "myVal");

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual("myVal", presenter.Show().ParameterName);
        }

        [TestMethod]
        public void Presenter_Reject_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithPropertyNameChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.Setup(v => v.NewPropertyName).Returns("MyProperty");
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual("MyProperty", presenter.Show().PropertyName);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.Setup(v => v.ImplementLetSetterType).Returns(true);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(true, presenter.Show().ImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.Setup(v => v.ImplementSetSetterType).Returns(true);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(true, presenter.Show().ImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetNotAllowedForPrimitiveTypes()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.IsSetterTypeChangeable, true);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            factory.Create().Show();

            Assert.AreEqual(false, view.Object.IsSetterTypeChangeable);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetNotAllowedForNonVariantNonPrimitiveType()
        {
            //Input
            const string inputCode =
@"Private fizz As Icon";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.IsSetterTypeChangeable, true);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            factory.Create().Show();

            Assert.AreEqual(false, view.Object.IsSetterTypeChangeable);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetAndSetAllowedForVariant()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.IsSetterTypeChangeable, true);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.IsSetterTypeChangeable);
        }

        [TestMethod]
        public void Presenter_Accept_DefaultCreateLetForPrimitiveTypes()
        {
            //Input
            const string inputCode =
@"Private fizz As Date";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.ImplementLetSetterType);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(true, model.ImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_DefaultCreateSetForNonVariantNonPrimitiveTypes()
        {
            //Input
            const string inputCode =
@"Private fizz As Icon";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.ImplementSetSetterType);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(true, model.ImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_DefaultCreateLetForVariant()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var view = new Mock<IEncapsulateFieldView>();
            view.SetupProperty(v => v.ImplementLetSetterType);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(parser.State, editor.Object, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(true, model.ImplementLetSetterType);
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IEncapsulateFieldPresenter>> SetupFactory(EncapsulateFieldModel model)
        {
            var presenter = new Mock<IEncapsulateFieldPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IEncapsulateFieldPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}