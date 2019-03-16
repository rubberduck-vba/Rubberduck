using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class ImplementInterfaceCommandTests : CodePaneRefactorCommandTestBase
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
            return new RefactorImplementInterfaceCommand(state, msgBox, rewritingManager, selectionService);
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
    }
}