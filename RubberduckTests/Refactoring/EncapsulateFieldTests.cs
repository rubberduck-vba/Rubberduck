using System;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class EncapsulateFieldTests : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidDeclarationType_Throws()
        {
            //Input
            const string inputCode =
                @"Public fizz As Integer";

            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, presenterAction, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_InvalidIdentifierSelected_Throws()
        {
            //Input
            const string inputCode =
                @"Public Function fizz() As Integer
End Function";
            var selection = new Selection(1, 19);

            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(inputCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementSet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name");
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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

            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementSet: true, implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var selection = new Selection(2, 2);

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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var selection = new Selection(3, 2);

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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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

            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(
                "Class1", 
                selection, 
                presenterAction, 
                null, 
                false, 
                ("Class1", codeClass1, ComponentType.ClassModule),
                ("Class2", codeClass2, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void EncapsulatePublicField_PassInTarget()
        {
            //Input
            const string inputCode =
                @"Private fizz As Integer";

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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, "fizz", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
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
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), Selection.Home);
                var selectionService = MockedSelectionService();
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(It.IsAny<EncapsulateFieldModel>()))
                    .Returns(() => null); // resolves ambiguous method overload

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

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

            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction = model => null;
            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(InvalidRefactoringModelException));
            Assert.AreEqual(inputCode, actualCode);
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
            var presenterAction = SetParameters("Name", implementLet: true);
            var actualCode = RefactoredCode(inputCode, "foo", DeclarationType.Variable, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        public void Refactoring_Puts_Code_In_Correct_Place()
        {
            //Input
            const string inputCode =
                @"Option Explicit

Public Foo As String";

            var selection = new Selection(3, 8, 3, 11);

            //Output
            const string expectedCode =
                @"Option Explicit

Private Foo As String

Public Property Get bar() As String
    bar = Foo
End Property

Public Property Let bar(ByVal value As String)
    Foo = value
End Property
";
            var presenterAction = SetParameters("bar", implementLet: true);
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        #region setup

        private Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(
            string propertyName,
            bool implementSet = false, 
            bool implementLet = false, 
            string parameterName = "value")
        {
            return model =>
            {
                model.PropertyName = propertyName;
                model.ParameterName = parameterName;
                model.ImplementLetSetterType = implementLet;
                model.ImplementSetSetterType = implementSet;
                return model;
            };
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider, uiDispatcherMock.Object);
        }

        #endregion
    }
}
