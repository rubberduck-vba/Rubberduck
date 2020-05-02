using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberToStdModulePreviewTests
    {
        private string OptionExplicitBlock = $"Option Explicit{Environment.NewLine}{Environment.NewLine}";

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
            var preview = RetrievePreviewsAfterUserInput(memberToMove, endpoints, source, createNewModule ? null : OptionExplicitBlock);

            Assert.IsTrue(OptionExplicitBlock.Trim().OccursOnce( preview.Destination));
            Assert.IsTrue("Public Function Foo(".OccursOnce(preview.Destination));
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
            var preview = RetrievePreviewsAfterUserInput(memberToMove, endpoints, source, createNewModule ? null : OptionExplicitBlock);

            StringAssert.Contains(OptionExplicitBlock.Trim(), preview.Destination);
            Assert.IsTrue("Public Sub Foo(".OccursOnce(preview.Destination));
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
            var preview = RetrievePreviewsAfterUserInput(memberToMove, endpoints, source, createNewModule ? null : OptionExplicitBlock);

            StringAssert.Contains(OptionExplicitBlock.Trim(), preview.Destination);
            Assert.IsTrue("Property Get TheValue(".OccursOnce(preview.Destination));
            Assert.IsTrue("Property Let TheValue(".OccursOnce(preview.Destination));
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
            var preview = RetrievePreviewsAfterUserInput(memberToMove, MoveEndpoints.StdToStd, source, OptionExplicitBlock);

            var changesBelowMarker = Rubberduck.Resources.RubberduckUI.MoveMember_NewContentBelowThisLine;
            var changesAboveMarker = Rubberduck.Resources.RubberduckUI.MoveMember_NewContentAboveThisLine;

            Assert.IsTrue(OptionExplicitBlock.Trim().OccursOnce(preview.Destination));
            Assert.IsTrue("Public Function Foo(".OccursOnce(preview.Destination));

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

            var destination =
$@"
Option Explicit

Private mTest As Long

Function Goo(arg1 As Long) As Long
    Goo = mTest * arg1
End Function
";
            var preview = RetrievePreviewsAfterUserInput(memberToMove, MoveEndpoints.StdToStd, source, destination);
            var changesBelowMarker = Rubberduck.Resources.RubberduckUI.MoveMember_NewContentBelowThisLine;
            var changesAboveMarker = Rubberduck.Resources.RubberduckUI.MoveMember_NewContentAboveThisLine;

            Assert.IsTrue(OptionExplicitBlock.Trim().OccursOnce(preview.Destination));
            Assert.IsTrue("Public Function Foo(".OccursOnce(preview.Destination));

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

            var destination =
$@"
Option Explicit

Private mTest As Long

Public mFizz As Long

Sub Fizz(arg1 As Long, ByRef outVal As Long)
    outVal = mTest * mFizz
End Sub
";
            var preview = RetrievePreviewsAfterUserInput(memberToMove, MoveEndpoints.StdToStd, source, destination);

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

            var destination =
$@"
Option Explicit

Private mTest As Long

Public mFizz As Long

Sub Fizz(arg1 As Long, ByRef outVal As Long)
    outVal = mTest * mFizz
End Sub
";
            var (state, rewritingManager) = CreateAndParseWithRewritingManager(MoveEndpoints.StdToStd, source, destination);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Function)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.Item1);

                var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);
                model.ChangeDestination(MoveEndpoints.StdToStd.DestinationModuleName(), ComponentType.StandardModule);

                //Show a preview that should have renamed a couple declarations

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var previewerFactory = resolver.Resolve<IMoveMemberRefactoringPreviewerFactory>();
                var preview = previewerFactory.Create(model.Destination).PreviewMove(model);

                StringAssert.Contains("mFizz1", preview);
                StringAssert.Contains("Public Function Fizz1", preview);

                model.ChangeDestination(null);

                preview = previewerFactory.Create(model.Destination).PreviewMove(model);
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
            var (state, rewritingManager) = CreateAndParseWithRewritingManager(MoveEndpoints.StdToStd, sourceContent, createNewModule ? null : OptionExplicitBlock);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.PropertyGet)
                                    .Single(declaration => declaration.IdentifierName == "TheValue");

                var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);
                if (!createNewModule)
                {
                    model.ChangeDestination(MoveEndpoints.StdToStd.DestinationModuleName(), ComponentType.StandardModule);
                }

                foreach (var moveable in model.MoveableMemberSets)
                {
                    moveable.IsSelected = false;
                }

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var previewerFactory = resolver.Resolve<IMoveMemberRefactoringPreviewerFactory>();

                var destination = previewerFactory.Create(model.Destination).PreviewMove(model);

                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NewContentBelowThisLine, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NewContentAboveThisLine, destination);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewEmptySelectionSetExistingDestinationModule()
        {
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
            var (state, rewritingManager) = CreateAndParseWithRewritingManager(MoveEndpoints.StdToStd, sourceContent, destinationContent);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.PropertyGet)
                                    .Single(declaration => declaration.IdentifierName == "TheValue");

                var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);
                model.ChangeDestination(MoveEndpoints.StdToStd.DestinationModuleName(), ComponentType.StandardModule);

                foreach (var moveable in model.MoveableMemberSets)
                {
                    moveable.IsSelected = false;
                }

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var previewerFactory = resolver.Resolve<IMoveMemberRefactoringPreviewerFactory>();

                var destination = previewerFactory.Create(model.Destination).PreviewMove(model);

                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NewContentBelowThisLine, destination);
                StringAssert.Contains(Rubberduck.Resources.RubberduckUI.MoveMember_NewContentAboveThisLine, destination);
            }

        }

        private (string Source, string Destination) RetrievePreviewsAfterUserInput((string declarationName, DeclarationType declarationType) memberToMove, MoveEndpoints endpoints, string sourceContent, string destinationContent)
        {
            var (state, rewritingManager) = CreateAndParseWithRewritingManager(endpoints, sourceContent, destinationContent);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(memberToMove.declarationType)
                                    .Single(declaration => declaration.IdentifierName == memberToMove.declarationName);

                var model = MoveMemberTestsResolver.CreateRefactoringModel(target, state);
                if (destinationContent != null)
                {
                    model.ChangeDestination(MoveEndpoints.StdToStd.DestinationModuleName(), ComponentType.StandardModule);
                }

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var previewerFactory = resolver.Resolve<IMoveMemberRefactoringPreviewerFactory>();

                var source = previewerFactory.Create(model.Source).PreviewMove(model);
                var destination = previewerFactory.Create(model.Destination).PreviewMove(model);

                return (source, destination);
            }
        }

        private (RubberduckParserState, IRewritingManager) CreateAndParseWithRewritingManager(MoveEndpoints endpoints, string sourceContent, string destinationContent)
        {
            var modules = endpoints.ToModulesTuples(sourceContent, destinationContent);
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return MockParser.CreateAndParseWithRewritingManager(vbe);
        }
    }
}
