using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class ImplementInterfaceCommandTests : RefactorCodePaneCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void ImplementInterface_CanExecute_ImplementsInterfaceNotSelected()
        {
            const string classCode = @"Implements IClass1
Dim b As Variant";
            var selection = new Selection(2,4);

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, string.Empty)
                .AddComponent("Class1", ComponentType.ClassModule, classCode, selection)
                .Build();
            var vbe = builder.AddProject(project).Build().Object;

            Assert.IsFalse(CanExecute(vbe));
        }

        [Category("Commands")]
        [Test]
        public void ImplementInterface_CanExecute_ImplementsInterfaceSelected()
        {
            const string classCode = @"Implements IClass1";
            var selection = Selection.Home;

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, string.Empty)
                .AddComponent("Class1", ComponentType.ClassModule, classCode, selection)
                .Build();
            var vbe = builder.AddProject(project).Build().Object;

            Assert.IsTrue(CanExecute(vbe));
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var msgBox = new Mock<IMessageBox>().Object;
            var addImplementationsBaseRefactoring = new AddInterfaceImplementationsRefactoringAction(rewritingManager, CreateCodeBuilder());
            var baseRefactoring = new ImplementInterfaceRefactoringAction(addImplementationsBaseRefactoring, rewritingManager);
            var refactoring = new ImplementInterfaceRefactoring(baseRefactoring, state, selectionService);
            var notifier = new ImplementInterfaceFailedNotifier(msgBox);
            return new RefactorImplementInterfaceCommand(refactoring, notifier, state, selectionService);
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string classCode = @"Implements IClass1";
            var selection = Selection.Home;

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, string.Empty)
                .AddComponent("Class1", ComponentType.ClassModule, classCode, selection)
                .Build();

            return builder.AddProject(project).Build().Object;
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
    }
}