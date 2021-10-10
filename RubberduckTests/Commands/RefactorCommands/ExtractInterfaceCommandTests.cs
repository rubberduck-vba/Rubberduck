using System;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class ExtractInterfaceCommandTests : RefactorCodePaneCommandTestBase
    {
        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_NoMembers()
        {
            const string input = @"Option Explicit";

            Assert.IsFalse(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_Proc_StdModule()
        {
            const string input =
                @"Sub foo()
End Sub";

            Assert.IsFalse(CanExecute(input, ComponentType.StandardModule));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_Field()
        {
            const string input = "Dim d As Boolean";

            Assert.IsFalse(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void CanExecuteNameCollision_ActiveCodePane_EmptyClass()
        {
            const string input = @"
Sub Foo()
End Sub
";
            var builder = new MockVbeBuilder();
            var proj1 = builder.ProjectBuilder("TestProj1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, input, Selection.Home)
                .Build();
            var proj2 = builder.ProjectBuilder("TestProj2", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty, Selection.Home)
                .Build();

            var vbe = builder
                .AddProject(proj1)
                .AddProject(proj2)
                .Build();

            vbe.Object.ActiveCodePane = proj1.Object.VBComponents[0].CodeModule.CodePane;
            if (string.IsNullOrEmpty(vbe.Object.ActiveCodePane.CodeModule.Content()))
            {
                Assert.Inconclusive("The active code pane should be the one with the method stub.");
            }

            Assert.IsTrue(CanExecute(vbe.Object));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_ClassWithMembers_SameNameAsClassWithMembers()
        {
            const string input =
                @"Sub foo()
End Sub";

            var builder = new MockVbeBuilder();
            var proj1 = builder.ProjectBuilder("TestProj1", ProjectProtection.Unprotected).AddComponent("Comp1", ComponentType.ClassModule, input, Selection.Home).Build();
            var proj2 = builder.ProjectBuilder("TestProj2", ProjectProtection.Unprotected).AddComponent("Comp1", ComponentType.ClassModule, string.Empty).Build();

            var vbe = builder
                .AddProject(proj1)
                .AddProject(proj2)
                .Build();

            vbe.Setup(s => s.ActiveCodePane).Returns(proj1.Object.VBComponents[0].CodeModule.CodePane);

            Assert.IsTrue(CanExecute(vbe.Object));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_Proc()
        {
            const string input =
                @"Sub foo()
End Sub";

            Assert.IsTrue(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_Function()
        {
            const string input =
                @"Function foo() As Integer
End Function";

            Assert.IsTrue(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_PropertyGet()
        {
            const string input =
                @"Property Get foo() As Boolean
End Property";

            Assert.IsTrue(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_PropertyLet()
        {
            const string input =
                @"Property Let foo(value)
End Property";

            Assert.IsTrue(CanExecute(input));
        }

        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_PropertySet()
        {
            const string input =
                @"Property Set foo(value)
End Property";

            Assert.IsTrue(CanExecute(input));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5693
        [Test]
        [Category("Commands")]
        [Category("Extract Interface")]
        public void ExtractInterface_CanExecute_ExistingImplements()
        {
            var input =
@"
Implements ITest

Property Set Foo(value)
End Property";

            Assert.IsTrue(CanExecute(input));
        }

        private bool CanExecute(string inputCode, ComponentType componentType = ComponentType.ClassModule)
        {
            return CanExecute(inputCode, Selection.Home, componentType);
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var factory = new Mock<IRefactoringPresenterFactory>().Object;
            var msgBox = new Mock<IMessageBox>().Object;
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            var addComponentService = TestAddComponentService(state.ProjectsProvider);
            var baseRefactoring = new ExtractInterfaceRefactoringAction(addImplementationsBaseRefactoring, state, state, rewritingManager, state.ProjectsProvider, addComponentService);
            var userInteraction = new RefactoringUserInteraction<IExtractInterfacePresenter, ExtractInterfaceModel>(factory, uiDispatcherMock.Object);
            var refactoring = new ExtractInterfaceRefactoring(baseRefactoring, state, userInteraction, selectionService, new Mock<IExtractInterfaceConflictFinderFactory>().Object, CreateCodeBuilder());
            var notifier = new ExtractInterfaceFailedNotifier(msgBox);
            return new RefactorExtractInterfaceCommand(refactoring, notifier, state, selectionService);
        }

        private static ICodeBuilder CreateCodeBuilder()
            => new CodeBuilder(new Indenter(null, CreateIndenterSettings));

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Property Let foo(value)
End Property";
            var selection = Selection.Home;
            return TestVbe(input, selection, ComponentType.ClassModule);
        }
    }
}