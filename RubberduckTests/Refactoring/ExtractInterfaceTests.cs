using System;
using System.Collections.ObjectModel;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractInterfaceTests : InteractiveRefactoringTestBase<IExtractInterfacePresenter, ExtractInterfaceModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProc()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                return model;
            };

            var actualCode = RefactoredCode("Class", selection, presenterAction, null, false, ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class"]);
            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_InvalidTargetType_Throws()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                return model;
            };

            var actualCode = RefactoredCode(
                "Module", 
                DeclarationType.ProceduralModule, 
                presenterAction, 
                typeof(InvalidDeclarationTypeException),
                ("Module", inputCode, ComponentType.StandardModule));
            Assert.AreEqual(inputCode, actualCode["Module"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NoValidTargetSelected_Throws()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                return model;
            };

            var actualCode = RefactoredCode(
                "Module",
                selection,
                presenterAction,
                typeof(NoDeclarationForSelectionException),
                false,
                ("Module", inputCode, ComponentType.StandardModule));
            Assert.AreEqual(inputCode, actualCode["Module"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFuncAndPropGetSetLet()
        {
            //Input
            const string inputCode = @"
Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode = @"
Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function IClass_Fizz(ByRef b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function

Private Property Get IClass_Buzz() As Variant
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Let IClass_Buzz(ByRef value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property

Private Property Set IClass_Buzz(ByRef value As Variant)
    Err.Raise 5 'TODO implement interface member
End Property
";

            const string expectedInterfaceCode =
                @"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

Public Property Get Buzz() As Variant
End Property

Public Property Let Buzz(ByRef value As Variant)
End Property

Public Property Set Buzz(ByRef value As Variant)
End Property

";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var interfaceMember in model.Members)
                {
                    interfaceMember.IsSelected = true;
                }

                return model;
            };

            var actualCode = RefactoredCode("Class", selection, presenterAction, null, false, ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class"]);
            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_ImplementProcAndFunc_IgnoreProperties()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(b) As Variant
End Function

Public Property Get Buzz()
End Property

Public Property Let Buzz(value)
End Property

Public Property Set Buzz(value)
End Property

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub

Private Function IClass_Fizz(ByRef b As Variant) As Variant
    Err.Raise 5 'TODO implement interface member
End Function
";

            const string expectedInterfaceCode =
                @"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Public Function Fizz(ByRef b As Variant) As Variant
End Function

";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                foreach (var interfaceMember in model.Members.Where(member => !member.FullMemberSignature.Contains("Property")))
                {
                    interfaceMember.IsSelected = true;
                }

                return model;
            };

            var actualCode = RefactoredCode("Class", selection, presenterAction, null, false, ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class"]);
            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            //Input
            const string inputCode =
                @"Public Fizz As Boolean";

            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _, selection);
            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var target  = state.DeclarationFinder.UserDeclarations(DeclarationType.ClassModule).First();

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, target);
                Assert.AreEqual(0, model.Members.Count);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NullPresenter_NoChanges()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out var component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                //SetupFactory
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(It.IsAny<ExtractInterfaceModel>())).Returns(value: null);

                var selectionService = MockedSelectionService();

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                Assert.AreEqual(1, vbe.Object.ActiveVBProject.VBComponents.Count());
                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_NullModel_NoChanges()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model => null;

            var actualCode = RefactoredCode("Class", selection, presenterAction, typeof(InvalidRefactoringModelException), false, ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(inputCode, actualCode["Class"]);
            Assert.AreEqual(1, actualCode.Count);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_PassTargetIn()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Expectation
            const string expectedCode =
                @"Implements IClass

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub IClass_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
    Err.Raise 5 'TODO implement interface member
End Sub
";

            const string expectedInterfaceCode =
                @"Option Explicit

Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

";
            Func<ExtractInterfaceModel, ExtractInterfaceModel> presenterAction = model =>
            {
                model.Members.ElementAt(0).IsSelected = true;
                model.Members = new ObservableCollection<InterfaceMember>(new[] { model.Members.ElementAt(0) }.ToList());

                return model;
            };

            var actualCode = RefactoredCode("Class", DeclarationType.ClassModule, presenterAction, null, ("Class", inputCode, ComponentType.ClassModule));
            Assert.AreEqual(expectedCode, actualCode["Class"]);
            var actualInterfaceCode = actualCode[actualCode.Keys.Single(componentName => !componentName.Equals("Class"))];
            Assert.AreEqual(expectedInterfaceCode, actualInterfaceCode);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            return new ExtractInterfaceRefactoring(state, state, factory, rewritingManager, selectionService, uiDispatcherMock.Object);
        }
    }
}