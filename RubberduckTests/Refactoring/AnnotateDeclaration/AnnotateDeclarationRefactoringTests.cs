using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring.AnnotateDeclaration
{
    [TestFixture]
    public class AnnotateDeclarationRefactoringTests : InteractiveRefactoringTestBase<IAnnotateDeclarationPresenter, AnnotateDeclarationModel>
    {
        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoring_InitialModel_TargetIsPassedInTarget()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            var targetName = model.Target.IdentifierName;

            Assert.AreEqual("TestModule", targetName);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoring_InitialModel_AnnotationNull()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            Assert.IsNull(model.Annotation);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoring_InitialModel_ArgumentsEmpty()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            var model = InitialModel(
                "TestModule",
                DeclarationType.ProceduralModule,
                ("TestModule", code, ComponentType.StandardModule));

            Assert.IsFalse(model.Arguments.Any());
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoring_InvalidTargetType_Throws()
        {
            const string code = @"Public Sub Foo()
myLabel: Debug.Print ""Label"";
End Sub";
            Assert.Throws<InvalidDeclarationTypeException>(() => 
                InitialModel(
                    "myLabel",
                    DeclarationType.LineLabel,
                    ("TestModule", code, ComponentType.StandardModule)));
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IAnnotateDeclarationPresenter, AnnotateDeclarationModel> userInteraction, 
            ISelectionService selectionService)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var attributesUpdater = new AttributesUpdater(state);
            var annotateDeclarationAction = new AnnotateDeclarationRefactoringAction(rewritingManager, annotationUpdater, attributesUpdater);

            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new AnnotateDeclarationRefactoring(annotateDeclarationAction, selectedDeclarationProvider, selectionService, userInteraction);
        }
    }
}