﻿using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Refactoring.DeleteDeclarations;
using RubberduckTests.Settings;
using System;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class MoveCloserToUsageCommandTests : RefactorCodePaneCommandTestBase
    {
        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_Field_NoReferences()
        {
            const string input =
                @"Dim d As Boolean";
            var selection = new Selection(2, 10, 2, 10);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_LocalVariable_NoReferences()
        {
            const string input =
                @"Property Get foo() As Boolean
    Dim d As Boolean
End Property";
            var selection = new Selection(2, 10, 2, 10);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_Const_NoReferences()
        {
            const string input =
                @"Private Const const_abc = 0";
            var selection = Selection.Home;

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_Field()
        {
            const string input =
                @"Dim d As Boolean
Sub Foo()
    d = True
End Sub";
            var selection = new Selection(1, 5, 1, 5);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_LocalVariable()
        {
            const string input =
                @"Property Get foo() As Boolean
    Dim d As Boolean
    d = True
End Property";
            var selection = new Selection(2, 10, 2, 10);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Test]
        public void MoveCloserToUsage_CanExecute_Const()
        {
            const string input =
                @"Private Const const_abc = 0
Sub Foo()
    Dim d As Integer
    d = const_abc
End Sub";
            var selection = new Selection(1, 17, 1, 17);

            Assert.IsTrue(CanExecute(input, selection));
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var factory = new Mock<IRefactoringPresenterFactory>().Object;
            var msgBox = new Mock<IMessageBox>().Object;

            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            var userInteraction = new RefactoringUserInteraction<IMoveCloserToUsagePresenter, MoveCloserToUsageModel>(factory, uiDispatcherMock.Object);

            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);

            var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteDeclarationsRefactoringAction>();

            var baseRefactoring = new MoveCloserToUsageRefactoringAction(deleteDeclarationRefactoringAction, rewritingManager);
            var refactoring = new MoveCloserToUsageRefactoring(baseRefactoring, state, selectionService, selectedDeclarationProvider, userInteraction);
            var notifier = new MoveCloserToUsageFailedNotifier(msgBox);
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            return new RefactorMoveCloserToUsageCommand(refactoring, notifier, state, selectionService, selectedDeclarationService);            
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Private Const const_abc = 0
Sub Foo()
    Dim d As Integer
    d = const_abc
End Sub";
            var selection = new Selection(1, 17, 1, 17);
            return TestVbe(input, selection);
        }
    }
}