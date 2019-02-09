using System.Windows.Forms;
using NUnit.Framework;
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
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class EncapsulateFieldTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = false,
                    ImplementSetSetterType = true,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = false,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
            
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);
                
                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = true,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var targetComponent = state.ProjectsProvider.Component(model.TargetDeclaration.QualifiedModuleName);
                var actualCode = targetComponent.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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
                .AddComponent("Class1", ComponentType.ClassModule, codeClass1, selection)
                .AddComponent("Class2", ComponentType.ClassModule, codeClass2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[0];
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var module1 = project.Object.VBComponents[0].CodeModule;
                var module2 = project.Object.VBComponents[1].CodeModule;

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode1 = module1.Content();
                var actualCode2 = module2.Content();

                Assert.AreEqual(expectedCode1, actualCode1);
                Assert.AreEqual(expectedCode2, actualCode2);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_PresenterIsNull()
        {
            //Input
            const string inputCode =
                @"Private fizz As Variant";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var vbeWrapper = vbe.Object;
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(It.IsAny<EncapsulateFieldModel>()))
                    .Returns(() => null); // resolves ambiguous method overload

                var refactoring = new EncapsulateFieldRefactoring(state, vbeWrapper, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor();

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulateField_ModelIsNull()
        {
            //Input
            const string inputCode =
                @"Private fizz As Variant";
            var selection = new Selection(1, 1);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //SetupFactory
                var factory = SetupFactory(null);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(inputCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_OptionExplicit_NotMoved()
        {
            //Input
            const string inputCode =
                @"Option Explicit

Public foo As String";
            var selection = new Selection(3, 9);

            //Expectation
            const string expectedCode =
                @"Option Explicit

Private foo As String

Public Property Get Name() As String
    Name = foo
End Property

Public Property Let Name(ByVal value As String)
    foo = value
End Property
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var model = new EncapsulateFieldModel(state, qualifiedSelection)
                {
                    ImplementLetSetterType = true,
                    ImplementSetSetterType = false,
                    //CanImplementLet = true,
                    ParameterName = "value",
                    PropertyName = "Name"
                };

                //SetupFactory
                var factory = SetupFactory(model);

                var refactoring = new EncapsulateFieldRefactoring(state, vbe.Object, CreateIndenter(vbe.Object), factory.Object, rewritingManager);
                refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        // FIXME this test is bollocks
        public void Presenter_ParameterlessTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15);

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var factory = SetupFactory(null);
                var presenter = factory.Object.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(null);

                Assert.AreEqual(null, presenter.Show());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        // FIXME the assumption of this test is a smart presenter. That's bollocks
        public void Presenter_NullTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
                @"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = Selection.Home;

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var codePane = vbe.Object.VBProjects[0].VBComponents[0].CodeModule.CodePane;
                codePane.Selection = selection;

                var factory = SetupFactory(null);

                var presenter = factory.Object.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(null);

                Assert.AreEqual(null, presenter.Show());
            }
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory> SetupFactory(EncapsulateFieldModel model)
        {
            var presenter = new Mock<IEncapsulateFieldPresenter>();
            var factory = new Mock<IRefactoringPresenterFactory>();
            factory.Setup(f => f.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(It.IsAny<EncapsulateFieldModel>()))
                .Callback(() => presenter.Setup(p => p.Show()).Returns(model))
                .Returns(presenter.Object);
            return factory;
        }

        private static IIndenter CreateIndenter(IVBE vbe)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
        #endregion
    }
}
