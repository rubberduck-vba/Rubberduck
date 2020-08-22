using System;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Commands.RefactorCommands
{
    [TestFixture]
    public class RemoveParametersCommandTests : RefactorCodePaneCommandTestBase
    {
        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Event_NoParams()
        {
            const string input =
                @"Public Event Foo()";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Proc_NoParams()
        {
            const string input =
                @"Sub foo()
End Sub";
            var selection = new Selection(1, 6, 1, 6);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Function_NoParams()
        {
            const string input =
                @"Function foo() As Integer
End Function";
            var selection = new Selection(1, 11, 1, 11);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertyGet_NoParams()
        {
            const string input =
                @"Property Get foo() As Boolean
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertyLet_OneParam()
        {
            const string input =
                @"Property Let foo(value)
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertySet_OneParam()
        {
            const string input =
                @"Property Set foo(value)
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsFalse(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Event_OneParam()
        {
            const string input =
                @"Public Event Foo(value)";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Proc_OneParam()
        {
            const string input =
                @"Sub foo(value)
End Sub";
            var selection = new Selection(1, 6, 1, 6);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_Function_OneParam()
        {
            const string input =
                @"Function foo(value) As Integer
End Function";
            var selection = new Selection(1, 11, 1, 11);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertyGet_OneParam()
        {
            const string input =
                @"Property Get foo(value) As Boolean
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertyLet_TwoParams()
        {
            const string input =
                @"Property Let foo(value1, value2)
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsTrue(CanExecute(input, selection));
        }

        [Category("Commands")]
        [Test]
        public void RemoveParameters_CanExecute_PropertySet_TwoParams()
        {
            const string input =
                @"Property Set foo(value1, value2)
End Property";
            var selection = new Selection(1, 16, 1, 16);

            Assert.IsTrue(CanExecute(input, selection));
        }

        protected override CommandBase TestCommand(IVBE vbe, RubberduckParserState state, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            var msgBox = new Mock<IMessageBox>().Object;
            var factory = new Mock<IRefactoringPresenterFactory>().Object;
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            var baseRefactoring = new RemoveParameterRefactoringAction(state, rewritingManager);
            var userInteraction = new RefactoringUserInteraction<IRemoveParametersPresenter, RemoveParametersModel>(factory, uiDispatcherMock.Object);
            var refactoring = new RemoveParametersRefactoring(baseRefactoring, state, userInteraction, selectionService, selectedDeclarationProvider);
            var notifier = new RemoveParameterFailedNotifier(msgBox);
            return new RefactorRemoveParametersCommand(refactoring, notifier, state, selectionService, selectedDeclarationProvider);
        }

        protected override IVBE SetupAllowingExecution()
        {
            const string input =
                @"Property Set foo(value1, value2)
End Property";
            var selection = new Selection(1, 16, 1, 16);
            return TestVbe(input, selection);
        }
    }
}