using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldIsOverMultipleLines()
        {
            //Input
            const string inputCode =
@"Public _
fizz _
As _
Integer";
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_WithSetter()
        {
            //Input
            const string inputCode =
@"Public fizz As Variant";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Variant

Public Property Get Name() As Variant
    Set Name = fizz
End Property

Public Property Set Name(ByVal value As Variant)
    Set fizz = value
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = false,
                ImplementSetSetterType = true,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_WithOnlyGetter()
        {
            //Input
            const string inputCode =
@"Public fizz As Variant";
            var selection = new Selection(1, 1);

            //Expectation
            const string expectedCode =
@"Private fizz As Variant

Public Property Get Name() As Variant
    Name = fizz
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = false,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_OtherFieldsInClass()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Public buzz As Boolean";
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveFirst()
        {
            //Input
            const string inputCode =
@"Public fizz, _
         buzz As Boolean, _
         bazz As Date";
            var selection = new Selection(1, 12);

            //Expectation
            const string expectedCode =
@"Public buzz As Boolean, _
         bazz As Date
Private fizz As Variant

Public Property Get Name() As Variant
    If IsObject(fizz) Then
        Set Name = fizz
    Else
        Name = fizz
    End If
End Property

Public Property Let Name(ByVal value As Variant)
    fizz = value
End Property

Public Property Set Name(ByVal value As Variant)
    Set fizz = value
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = true,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveSecond()
        {
            //Input
            const string inputCode =
@"Public fizz, _
buzz As Boolean, _
bazz As Date";
            var selection = new Selection(2, 12);

            //Expectation
            const string expectedCode =
@"Public fizz, _
bazz As Date
Private buzz As Boolean

Public Property Get Name() As Boolean
    Name = buzz
End Property

Public Property Let Name(ByVal value As Boolean)
    buzz = value
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_FieldDeclarationHasMultipleFields_MoveLast()
        {
            //Input
            const string inputCode =
@"Public fizz, _
buzz As Boolean, _
bazz As Date";
            var selection = new Selection(3, 12);

            //Expectation
            const string expectedCode =
@"Public fizz, _
buzz As Boolean
Private bazz As Date

Public Property Get Name() As Date
    Name = bazz
End Property

Public Property Let Name(ByVal value As Date)
    bazz = value
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulatePrivateField()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(model.TargetDeclaration);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void GivenReferencedPublicField_UpdatesReferenceToNewProperty()
        {
            //Input
            const string codeClass1 =
@"Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            const string codeClass2 =
@"Sub Foo()
    Dim c As Class1
    c.fizz = 0
    Bar c.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var selection = new Selection(1, 1);

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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, codeClass1)
                .AddComponent("Class2", ComponentType.ClassModule, codeClass2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var state = MockParser.CreateAndParse(vbe.Object);
            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var actualCode1 = module1.Content();
            var actualCode2 = module2.Content();

            Assert.AreEqual(expectedCode1, actualCode1);
            Assert.AreEqual(expectedCode2, actualCode2);

            var rewriter1 = state.GetRewriter(module1.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = state.GetRewriter(module2.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
        }

        [TestMethod]
        public void EncapsulatePublicField_PassInTarget()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";
            var selection = new Selection(1, 1);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new EncapsulateFieldModel(state, qualifiedSelection)
            {
                ImplementLetSetterType = true,
                ImplementSetSetterType = false,
                CanImplementLet = true,
                ParameterName = "value",
                PropertyName = "Name"
            };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulateField_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var vbeWrapper = vbe.Object;
            var factory = new EncapsulateFieldPresenterFactory(vbeWrapper, state, null);

            var refactoring = new EncapsulateFieldRefactoring(vbeWrapper, CreateIndenter(vbe.Object), factory);
            refactoring.Refactor();

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
        }

        [TestMethod]
        public void EncapsulateField_ModelIsNull()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 1);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //SetupFactory
            var factory = SetupFactory(null);

            var refactoring = new EncapsulateFieldRefactoring(vbe.Object, CreateIndenter(vbe.Object), factory.Object);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
        }

        [TestMethod]
        public void GivenNullActiveCodePane_FactoryReturnsNullPresenter()
        {
            //Input
            const string inputCode =
@"Private fizz As Integer";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            vbe.Object.ActiveCodePane = null;
            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, null);
            var actual = factory.Create();

            Assert.IsNull(actual);
        }

        [TestMethod]
        public void Presenter_ParameterlessTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, null);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var codePane = vbe.Object.VBProjects[0].VBComponents[0].CodeModule.CodePane;
            codePane.Selection = selection;

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithParameterNameChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var view = new Mock<IRefactoringDialog<EncapsulateFieldViewModel>>();
            view.Setup(v => v.DialogResult).Returns(DialogResult.OK);
            view.SetupGet(v => v.ViewModel).Returns(new EncapsulateFieldViewModel(state, null) {ParameterName = "myVal"});

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);

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
            var selection = new Selection(1, 15);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var view = new Mock<IRefactoringDialog<EncapsulateFieldViewModel>>();
            view.Setup(v => v.DialogResult).Returns(DialogResult.Cancel);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        //NOTE: The tests below are commented out pending some sort of refactoring that enables them
        //to actually *test* something.  Currently, all of the behavior the tests are looking for is
        //being mocked.

        /*
        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetChanged()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementLetSetterType, true);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);

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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementSetSetterType, true);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);

            var presenter = factory.Create();

            Assert.AreEqual(true, presenter.Show().ImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetAllowedForPrimitiveTypes_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementLetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.CanImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetNotAllowedForPrimitiveTypes_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementSetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(false, view.Object.CanImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetAllowedForNonVariantNonPrimitiveTypes_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Icon";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementSetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.CanImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetNotAllowedForNonVariantNonPrimitiveType_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Icon";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementLetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(false, view.Object.CanImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetAllowedForVariant_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementLetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.CanImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetAllowedForVariant_NoReferences()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.CanImplementSetSetterType, true);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.CanImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetRequiredForPrimitiveTypes_References()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean
Sub foo()
    fizz = True
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementLetSetterType, false);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.MustImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetRequiredForNonVariantNonPrimitiveTypes_References()
        {
            //Input
            const string inputCode =
@"Private fizz As Class1
Sub foo()
    Set fizz = New Class1
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementSetSetterType, false);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.MustImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementLetRequiredForNonSetVariant_References()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant
Sub Foo()
    fizz = True
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementLetSetterType, false);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.MustImplementLetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithImplementSetRequiredForSetVariant_References()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant
Sub foo()
    Set fizz = New Class1
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.SetupProperty(v => v.MustImplementSetSetterType, false);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            factory.Create().Show();

            Assert.AreEqual(true, view.Object.MustImplementSetSetterType);
        }


        [TestMethod]
        public void Presenter_Accept_DefaultCreateGetOnly_PrimitiveType_NoReference()
        {
            //Input
            const string inputCode =
@"Private fizz As Date";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(false, model.ImplementLetSetterType);
            Assert.AreEqual(false, model.ImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_DefaultCreateGetOnly_NonPrimitiveTypeNonVariant_NoReference()
        {
            //Input
            const string inputCode =
@"Private fizz As Icon";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(false, model.ImplementLetSetterType);
            Assert.AreEqual(false, model.ImplementSetSetterType);
        }

        [TestMethod]
        public void Presenter_Accept_DefaultCreateGetOnly_Variant_NoReference()
        {
            //Input
            const string inputCode =
@"Private fizz As Variant";
            var selection = new Selection(1, 15, 1, 15);

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (state.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var view = new Mock<IEncapsulateFieldDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new EncapsulateFieldPresenterFactory(vbe.Object, state, view.Object);
            var model = factory.Create().Show();

            Assert.AreEqual(false, model.ImplementLetSetterType);
            Assert.AreEqual(false, model.ImplementSetSetterType);
        }
        */
        #region setup
        private static Mock<IRefactoringPresenterFactory<IEncapsulateFieldPresenter>> SetupFactory(EncapsulateFieldModel model)
        {
            var presenter = new Mock<IEncapsulateFieldPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IEncapsulateFieldPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        private static IIndenter CreateIndenter(IVBE vbe)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
        #endregion
    }
}
