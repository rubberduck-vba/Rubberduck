using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public abstract class RefactorCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void RefactorCommand_CanExecute_NonReadyState()
        {
            var vbe = SetupAllowingExecution();
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations, CancellationToken.None);

                var selectionService = MockedSelectionService(vbe);
                var testCommand = TestCommand(vbe, state, rewritingManager, selectionService);
                Assert.IsFalse(testCommand.CanExecute(null));
            }
        }

        protected bool CanExecute(string inputCode, Selection? selection, ComponentType componentType = ComponentType.StandardModule)
        {
            var vbe = TestVbe(inputCode, selection ?? default, componentType);
            if (selection == null)
            {
                vbe.ActiveCodePane = null;
            }
            return CanExecute(vbe);
        }

        protected IVBE TestVbe(string inputCode, Selection selection, ComponentType componentType = ComponentType.StandardModule)
        {
            return MockVbeBuilder.BuildFromSingleModule(inputCode, componentType, out _, selection).Object;
        }

        protected bool CanExecute(IVBE vbe)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var selectionService = MockedSelectionService(vbe);
                var testCommand = TestCommand(vbe, state, rewritingManager, selectionService);
                return testCommand.CanExecute(null);
            }
        }


        protected abstract CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService);
        protected abstract IVBE SetupAllowingExecution();

        protected ISelectionService MockedSelectionService(IVBE vbe)
        {
            var selectionServiceMock = new Mock<ISelectionService>();
            var activeSelection = vbe.ActiveCodePane?.GetQualifiedSelection();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            return selectionServiceMock.Object;
        }
    }
}