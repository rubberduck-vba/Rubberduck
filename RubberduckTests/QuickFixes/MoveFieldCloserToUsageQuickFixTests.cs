using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete.Refactoring;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using RubberduckTests.Refactoring.DeleteDeclarations;
using Castle.Windsor;
using Rubberduck.Parsing.UIContext;
using Rubberduck.UI;
using System;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class MoveFieldCloserToUsageQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MoveFieldCloserToUsage_QuickFixWorks()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Static bar As String
    bar = ""test""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineIfStatement()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If bar = ""test"" Then Baz Else Foobar
End Sub

Private Sub Baz()
End Sub

Private Sub FooBar()
End Sub
";

            const string expectedCode =
                @"Public Sub Foo()
    Static bar As String
    If bar = ""test"" Then Baz Else Foobar
End Sub

Private Sub Baz()
End Sub

Private Sub FooBar()
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineThenStatement()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If True Then bar = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Static bar As String
    If True Then bar = ""test""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineElseStatement()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If True Then Else bar = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Static bar As String
    If True Then Else bar = ""test""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        private string ApplyQuickFixToFirstInspectionResult(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var resultToFix = inspectionResults.First();
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var selectionService = MockedSelectionService();
                var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);

                var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                    .Resolve<DeleteDeclarationsRefactoringAction>();

                var presenterMock = new Mock<IMoveCloserToUsagePresenter>();
                var factory = MockedRefractoringPresenterFactory(presenterMock);
                var uiDispatcher = MockedUiDispatcher();
                var userInteraction = new RefactoringUserInteraction<IMoveCloserToUsagePresenter, MoveCloserToUsageModel>(factory, uiDispatcher);

                var baseRefactoring = new MoveCloserToUsageRefactoringAction(deleteDeclarationRefactoringAction, rewritingManager);
                var refactoring = new MoveCloserToUsageRefactoring(baseRefactoring, state, selectionService, selectedDeclarationProvider, userInteraction);
                var quickFix = new MoveFieldCloserToUsageQuickFix(refactoring);

                quickFix.Fix(resultToFix, rewriteSession);

                return component.CodeModule.Content();
            }
        }

        private static ISelectionService MockedSelectionService()
        {
            QualifiedSelection? activeSelection = null;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }

        private static IRefactoringPresenterFactory MockedRefractoringPresenterFactory(Mock<IMoveCloserToUsagePresenter> presenterMock )
        {
            var factoryMock = new Mock<IRefactoringPresenterFactory>();
            factoryMock.Setup(f => f.Create<IMoveCloserToUsagePresenter, MoveCloserToUsageModel>(It.IsAny<MoveCloserToUsageModel>()))
                .Callback((MoveCloserToUsageModel model) => presenterMock.Setup(p => p.Show()).Returns(() => model))
                .Returns(presenterMock.Object);
            return factoryMock.Object;
        }

        private static IUiDispatcher MockedUiDispatcher()
        {
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            return uiDispatcherMock.Object;
        }
    }
}
