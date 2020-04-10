using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberToStdModulePreviewTests : InteractiveRefactoringTestBase<IMoveMemberPresenter, MoveMemberModel>
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

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);

            Assert.IsTrue(OccursOnce("Option Explicit", preview.Destination));
            Assert.IsTrue(OccursOnce("Public Function Foo(", preview.Destination));
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

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview.Destination);
            Assert.IsTrue(OccursOnce("Public Sub Foo(", preview.Destination));
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

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview.Destination);
            Assert.IsTrue(OccursOnce("Property Get TheValue(", preview.Destination));
            Assert.IsTrue(OccursOnce("Property Let TheValue(", preview.Destination));
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewCommentsSurroundNewContentNewModule()
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
            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove)
            {
                CreateNewModule = true
            };

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);
            var changesBelowMarker = Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine;
            var changesAboveMarker = Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine;

            Assert.IsTrue(OccursOnce("Option Explicit", preview.Destination));
            Assert.IsTrue(OccursOnce("Public Function Foo(", preview.Destination));

            StringAssert.Contains(changesBelowMarker, preview.Destination);
            StringAssert.Contains(changesAboveMarker, preview.Destination);
            Assert.IsTrue(preview.Destination.IndexOf("Option") < preview.Destination.IndexOf("Foo"));
            Assert.IsTrue(preview.Destination.IndexOf(changesBelowMarker) < preview.Destination.IndexOf("Foo"));
            Assert.IsTrue(preview.Destination.IndexOf(changesAboveMarker) > preview.Destination.IndexOf("Foo"));
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewCommentsSurroundNewContentExistingDestination()
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source =
$@"
Option Explicit

Function Foo(arg1 As Long) As Long
    Foo = arg1 + 2
End Function
";

            var existingContent =
$@"
Option Explicit

Private mTest As Long

Function Goo(arg1 As Long) As Long
    Goo = mTest * arg1
End Function
";


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove);
            moveDefinition.SetEndpointContent(source, existingContent);

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);
            var changesBelowMarker = Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine;
            var changesAboveMarker = Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine;

            Assert.IsTrue(OccursOnce("Option Explicit", preview.Destination));
            Assert.IsTrue(OccursOnce("Public Function Foo(", preview.Destination));

            StringAssert.Contains(changesBelowMarker, preview.Destination);
            StringAssert.Contains(changesAboveMarker, preview.Destination);
            Assert.IsTrue(preview.Destination.IndexOf("Option") < preview.Destination.IndexOf("Foo"));
            Assert.IsTrue(preview.Destination.IndexOf(changesBelowMarker) < preview.Destination.IndexOf("Foo"));
            Assert.IsTrue(preview.Destination.IndexOf(changesAboveMarker) > preview.Destination.IndexOf("Foo"));

            Assert.IsTrue(preview.Destination.IndexOf(changesBelowMarker) > preview.Destination.IndexOf("mTest"));
            Assert.IsTrue(preview.Destination.IndexOf(changesBelowMarker) < preview.Destination.IndexOf("Goo"));
            Assert.IsTrue(preview.Destination.IndexOf(changesAboveMarker) < preview.Destination.IndexOf("Goo"));

            var threeNewLines = string.Concat(Enumerable.Repeat(Environment.NewLine, 3));
            StringAssert.DoesNotContain(threeNewLines, preview.Source);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void IncludesRenamingActions()
        {
            var memberToMove = ("Fizz", DeclarationType.Function);
            var source =
$@"
Option Explicit

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function
";

            var existingContent =
$@"
Option Explicit

Private mTest As Long

Public mFizz As Long

Sub Fizz(arg1 As Long, ByRef outVal As Long)
    outVal = mTest * mFizz
End Sub
";


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove);
            moveDefinition.SetEndpointContent(source, existingContent);

            var preview = RetrievePreviewsAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("mFizz1", preview.Destination);
            StringAssert.Contains("Public Function Fizz1", preview.Destination);
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ResetsRenamingActions()
        {
            var memberToMove = ("Fizz", DeclarationType.Function);
            var source =
$@"
Option Explicit

Private mFizz As Long

Public Function Fizz() As Long
    Fizz = mFizz
End Function
";

            var existingContent =
$@"
Option Explicit

Private mTest As Long

Public mFizz As Long

Sub Fizz(arg1 As Long, ByRef outVal As Long)
    outVal = mTest * mFizz
End Sub
";


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove);
            moveDefinition.SetEndpointContent(source, existingContent);

            var preview = string.Empty;
            var vbe = BuildVBEStub(moveDefinition, source);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                moveDefinition.RewritingManager = rewritingManager;

                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Function)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.Item1);

               var  model = moveDefinition.ModelBuilder(state);

                //Show a preview that should have renamed a couple declarations
                InitializeThePreviewerFactory(model, state, rewritingManager);
                model.TryGetPreview(model.Destination, out preview);
                StringAssert.Contains("mFizz1", preview);
                StringAssert.Contains("Public Function Fizz1", preview);

                model.ChangeDestination(null);

                //The previously renamed no longer need renaming
                model.TryGetPreview(model.Destination, out preview);
                StringAssert.DoesNotContain("mFizz1", preview);
                StringAssert.DoesNotContain("Public Function Fizz1", preview);
            }
        }

        [TestCase(false)] //Null destination case
        [TestCase(true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewEmptySelectionSetNullOrNewDestinationModule(bool createNewModule)
        {
            var memberToMove = ("TheValue", DeclarationType.PropertyGet);
            var sourceContent =
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove)
            {
                CreateNewModule = createNewModule
            };

            var vbe = BuildVBEStub(moveDefinition, sourceContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                moveDefinition.RewritingManager = rewritingManager;

                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.PropertyGet)
                                    .Single(declaration => declaration.IdentifierName == "TheValue");

                var model = moveDefinition.ModelBuilder(state);

                foreach (var moveable in model.MoveableMembers)
                {
                    moveable.IsSelected = false;
                }

                if (!createNewModule)
                {
                    model.ChangeDestination(null);
                }

                InitializeThePreviewerFactory(model, state, rewritingManager);
                model.TryGetPreview(model.Source, out var source);
                model.TryGetPreview(model.Destination, out var destination);

                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine, destination);
            }

        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewEmptySelectionSetExistingDestinationModule()
        {
            var memberToMove = ("TheValue", DeclarationType.PropertyGet);
            var sourceContent =
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

            var destinationContent =
$@"
Option Explicit

Public Sub Fizz()
End Sub
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove);
            moveDefinition.SetEndpointContent(sourceContent, destinationContent);

            var vbe = BuildVBEStub(moveDefinition, sourceContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                moveDefinition.RewritingManager = rewritingManager;

                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.PropertyGet)
                                    .Single(declaration => declaration.IdentifierName == "TheValue");

                var model = moveDefinition.ModelBuilder(state);

                foreach (var moveable in model.MoveableMembers)
                {
                    moveable.IsSelected = false;
                }
                InitializeThePreviewerFactory(model, state, rewritingManager);
                model.TryGetPreview(model.Source, out var source);
                model.TryGetPreview(model.Destination, out var destination);

                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine, destination);
            }

        }

        private (string Source, string Destination) RetrievePreviewsAfterUserInput(TestMoveDefinition moveDefinition, string sourceContent, (string declarationName, DeclarationType declarationType) memberToMove)
        {
            var vbe = BuildVBEStub(moveDefinition, sourceContent);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                moveDefinition.RewritingManager = rewritingManager;

                var target = state.DeclarationFinder.DeclarationsWithType(memberToMove.declarationType)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.declarationName);

                var model = moveDefinition.ModelBuilder(state);

                InitializeThePreviewerFactory(model, state, rewritingManager);

                model.TryGetPreview(model.Source, out var source);
                model.TryGetPreview(model.Destination, out var destination);
                return (source, destination);
            }
        }

        private static void InitializeThePreviewerFactory(MoveMemberModel model, RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var tdi = new MoveMemberTestsDI(state, rewritingManager);
            var previewer = tdi.Resolve<IMoveMemberRefactoringPreviewerFactory>();
            model.PreviewerFactory = previewer;
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

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction, ISelectionService selectionService)
        {
            return MoveMemberTestSupport.CreateRefactoring(rewritingManager, state, userInteraction, selectionService);
        }
    }
}
