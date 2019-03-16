using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class IntroduceFieldCommandTests : CodePaneRefactorCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void IntroduceField_CanExecute_Field()
        {
            const string input =
                @"Dim d As Boolean";
            var selection = Selection.Home;

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void IntroduceField_CanExecute_LocalVariable()
        {
            const string input =
                @"Property Get foo() As Boolean
    Dim d As Boolean
End Property";
            var selection = new Selection(2, 10, 2, 10);
            Assert.IsTrue(CanExecute(input, selection));
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var msgBox = new Mock<IMessageBox>().Object;
            return new RefactorIntroduceFieldCommand(state, msgBox, rewritingManager, selectionService);
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Property Get foo() As Boolean
    Dim d As Boolean
End Property";
            var selection = new Selection(2, 10, 2, 10);
            return TestVbe(input, selection);
        }
    }
}