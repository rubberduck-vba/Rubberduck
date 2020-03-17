using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberPreviewTests : InteractiveRefactoringTestBase<IMoveMemberPresenter, MoveMemberModel>
    {
        [TestCase(MoveEndpoints.StdToStd, true)]
        [TestCase(MoveEndpoints.ClassToStd, true)]
        [TestCase(MoveEndpoints.FormToStd, true)]
        [TestCase(MoveEndpoints.StdToStd, false)]
        [TestCase(MoveEndpoints.ClassToStd, false)]
        [TestCase(MoveEndpoints.FormToStd, false)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContentFunction(MoveEndpoints endpoints, bool createNewModule)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source =
$@"
Option Explicit

Function Foo(arg1 As Long) As Long
    Const localConst As Long = 5
    Dim local As Long
    local = 6
    Foo = localConst + localVar + arg1
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, memberToMove)
            {
                CreateNewModule = createNewModule
            };

            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            Assert.IsTrue(OccursOnce("Option Explicit", preview));
            Assert.IsTrue(OccursOnce("Public Function Foo(", preview));
        }

        [TestCase(MoveEndpoints.StdToStd, true)]
        [TestCase(MoveEndpoints.ClassToStd, true)]
        [TestCase(MoveEndpoints.FormToStd, true)]
        [TestCase(MoveEndpoints.StdToStd, false)]
        [TestCase(MoveEndpoints.ClassToStd, false)]
        [TestCase(MoveEndpoints.FormToStd, false)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContentProcedure(MoveEndpoints endpoints, bool createNewModule)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
            var source =
$@"
Option Explicit

Sub Foo(ByVal arg1 As Long, ByRef result As Long)
    result = 10 * arg1
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, memberToMove)
            {
                CreateNewModule = createNewModule
            };

            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview);
            Assert.IsTrue(OccursOnce("Public Sub Foo(", preview));
        }


        [TestCase(MoveEndpoints.StdToStd, true)]
        [TestCase(MoveEndpoints.ClassToStd, true)]
        [TestCase(MoveEndpoints.FormToStd, true)]
        [TestCase(MoveEndpoints.StdToStd, false)]
        [TestCase(MoveEndpoints.ClassToStd, false)]
        [TestCase(MoveEndpoints.FormToStd, false)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContentProperties(MoveEndpoints endpoints, bool createNewModule)
        {
            var memberToMove = ("TheValue", DeclarationType.PropertyGet);
            var source =
$@"
Option Explicit


Private mTheValue As Long

Public Property Get TheValue() As Long
    TheValue = mTheValue
End Property

Public Property Let TheValue(ByVal value As Long)
    mTheValue = value
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, memberToMove)
            {
                CreateNewModule = createNewModule
            };

            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview);
            Assert.IsTrue(OccursOnce("Property Get TheValue(", preview));
            Assert.IsTrue(OccursOnce("Property Let TheValue(", preview));
        }

        private string RetrievePreviewAfterUserInput(TestMoveDefinition moveDefinition, string sourceContent, (string declarationName, DeclarationType declarationType) memberToMove)
        {
            MoveMemberModel PresenterAdjustment(MoveMemberModel model)
            {
                return moveDefinition.ModelBuilder(model.DeclarationFinderProvider);
            }

            var vbe = BuildVBEStub(moveDefinition, sourceContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(memberToMove.declarationType)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.declarationName);

                var refactoring = TestRefactoring(rewritingManager, state, PresenterAdjustment);
                if (refactoring is IMoveMemberRefactoringTestAccess testAccess)
                {
                    var model = testAccess.TestUserInteractionOnly(target, PresenterAdjustment);
                    return testAccess.PreviewModuleContent(model, PreviewModule.Destination);
                }
                throw new InvalidCastException();
            }
        }

        private static IVBE BuildVBEStub(TestMoveDefinition moveDefinition, string sourceContent)
        {
            if (moveDefinition.CreateNewModule)
            {
                moveDefinition.SetEndpointContent(sourceContent);
                return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple)).Object;
            }
            moveDefinition.SetEndpointContent(sourceContent, null);
            return MockVbeBuilder.BuildFromModules(moveDefinition.ModuleDefinitions.Select(tc => tc.AsTuple)).Object;
        }

        private static bool OccursOnce(string toFind, string content)
        {
            var firstIdx = content.IndexOf(toFind);
            var lastIdx = content.LastIndexOf(toFind);
            return firstIdx == lastIdx;
        }

        private static IAddComponentService TestAddComponentService(IProjectsProvider projectsProvider)
        {
            var sourceCodeHandler = new CodeModuleComponentSourceCodeHandler();
            return new AddComponentService(projectsProvider, sourceCodeHandler, sourceCodeHandler);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());

            var addComponentService = TestAddComponentService(state?.ProjectsProvider);
            var existingDestinationModuleRefactoring = new MoveMemberToExistingModuleRefactoring(state, rewritingManager);
            var newDestinationModuleRefactoring = new MoveMemberToNewModuleRefactoring(existingDestinationModuleRefactoring, state, rewritingManager, addComponentService);
            var refactoringAction = new MoveMemberRefactoringAction(newDestinationModuleRefactoring, existingDestinationModuleRefactoring);
            return new MoveMemberRefactoring(refactoringAction, state, factory, rewritingManager, selectionService, selectedDeclarationService, uiDispatcherMock.Object);
        }
    }
}
