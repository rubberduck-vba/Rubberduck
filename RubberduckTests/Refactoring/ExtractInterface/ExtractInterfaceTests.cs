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
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SourceCodeHandling;
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

'@Interface

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
        public void ExtractInterfaceRefactoring_IgnoresField()
        {
            //Input
            const string inputCode =
                @"Public Fizz As Boolean";

            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _, selection);
            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var target  = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .OfType<ClassModuleDeclaration>()
                    .First();

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, target, new CodeBuilder());
                Assert.AreEqual(0, model.Members.Count);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_DefaultsToPublicInterfaceForExposedImplementingClass()
        {
            //Input
            const string inputCode =
                @"Attribute VB_Exposed = True

Public Sub Foo
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var target = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .OfType<ClassModuleDeclaration>()
                    .First();

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, target, new CodeBuilder());
                Assert.AreEqual(ClassInstancing.Public, model.InterfaceInstancing);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ExtractInterfaceRefactoring_DefaultsToPrivateInterfaceForNonExposedImplementingClass()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _, selection);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var target = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .OfType<ClassModuleDeclaration>()
                    .First();

                //Specify Params to remove
                var model = new ExtractInterfaceModel(state, target, new CodeBuilder());
                Assert.AreEqual(ClassInstancing.Private, model.InterfaceInstancing);
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

'@Interface

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

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IExtractInterfacePresenter, ExtractInterfaceModel> userInteraction, 
            ISelectionService selectionService)
        {
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, new CodeBuilder());
            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            var baseRefactoring = new ExtractInterfaceRefactoringAction(addImplementationsBaseRefactoring, state, state, rewritingManager, state?.ProjectsProvider, addComponentService);
            return new ExtractInterfaceRefactoring(baseRefactoring, state, userInteraction, selectionService, new CodeBuilder());
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }
    }
}