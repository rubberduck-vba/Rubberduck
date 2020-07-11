using System;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class CodePaneRefactorMoveToFolderCommandTests : RefactorCodePaneCommandTestBase
    {
        //The only relevant test is in the base class.

        protected override CommandBase TestCommand(
            IVBE vbe, 
            RubberduckParserState state, 
            IRewritingManager rewritingManager,
            ISelectionService selectionService)
        {
            var factory = new Mock<IRefactoringPresenterFactory>().Object;
            var msgBox = new Mock<IMessageBox>().Object;

            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            var userInteraction = new RefactoringUserInteraction<IMoveMultipleToFolderPresenter, MoveMultipleToFolderModel>(factory, uiDispatcherMock.Object);

            var annotationUpdater = new AnnotationUpdater(state);
            var moveToFolderAction = new MoveToFolderRefactoringAction(rewritingManager, annotationUpdater);
            var moveMultipleToFolderAction = new MoveMultipleToFolderRefactoringAction(rewritingManager, moveToFolderAction);

            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);

            var refactoring = new MoveToFolderRefactoring(moveMultipleToFolderAction, selectedDeclarationProvider, selectionService, userInteraction, state);
            var notifier = new MoveToFolderRefactoringFailedNotifier(msgBox);

            return new CodePaneRefactorMoveToFolderCommand(refactoring, notifier, selectionService, state, selectedDeclarationProvider);
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Public Sub Foo()
End Sub";
            var selection = Selection.Home;
            return TestVbe(input, selection, ComponentType.ClassModule);
        }
    }
}