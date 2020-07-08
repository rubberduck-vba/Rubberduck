using System;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete.Refactoring;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnusedParameterQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        [Apartment(ApartmentState.STA)]
        public void GivenPrivateSub_DefaultQuickFixRemovesParameter()
        {
            const string inputCode = @"
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"
Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                var selectionService = MockedSelectionService();

                var factory = new Mock<IRefactoringPresenterFactory>().Object;
                var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
                var uiDispatcherMock = new Mock<IUiDispatcher>();
                uiDispatcherMock
                    .Setup(m => m.Invoke(It.IsAny<Action>()))
                    .Callback((Action action) => action.Invoke());
                var baseRefactoring = new RemoveParameterRefactoringAction(state, rewritingManager);
                var userInteraction = new RefactoringUserInteraction<IRemoveParametersPresenter, RemoveParametersModel>(factory, uiDispatcherMock.Object);
                var refactoring = new RemoveParametersRefactoring(baseRefactoring, state, userInteraction, selectionService, selectedDeclarationProvider);
                new RemoveUnusedParameterQuickFix(refactoring)
                    .Fix(inspectionResults.First(), rewriteSession);
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
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
    }
}
