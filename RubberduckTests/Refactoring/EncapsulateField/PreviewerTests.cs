using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldPreviewerTests : EncapsulateFieldInteractiveRefactoringTest
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [TestCase(EncapsulateFieldStrategy.UseBackingFields)]
        [TestCase(EncapsulateFieldStrategy.ConvertFieldsToUDTMembers)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldPreviewProvider))]
        public void Preview_EditPropertyIdentifier(EncapsulateFieldStrategy strategy)
        {
            var inputCode =
$@"Option Explicit

Public mTest As Long
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.StandardModule, out _);
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var resolver = Support.SetupResolver(state, rewritingManager);

                var target = state.DeclarationFinder.MatchName("mTest").First();
                var modelfactory = resolver.Resolve<IEncapsulateFieldModelFactory>();
                var model = modelfactory.Create(target);

                model.EncapsulateFieldStrategy = strategy;
                var field = model["mTest"];
                field.PropertyIdentifier = "ATest";

                var previewProvider = resolver.Resolve<EncapsulateFieldPreviewProvider>();

                var firstPreview = previewProvider.Preview(model);
                StringAssert.Contains("Property Get ATest", firstPreview);

                field.PropertyIdentifier = "BTest";
                var secondPreview = previewProvider.Preview(model);
                StringAssert.Contains("Property Get BTest", secondPreview);
                StringAssert.DoesNotContain("Property Get ATest", secondPreview);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldPreviewProvider))]
        public void PreviewWrapMember_EditPropertyIdentifier()
        {
            var inputCode =
$@"Option Explicit

Private Type T{MockVbeBuilder.TestModuleName}
    FirstValue As Long
End Type

Private Type B{MockVbeBuilder.TestModuleName}
    FirstValue As Long
End Type

Public mTest As Long

Private tType As T{MockVbeBuilder.TestModuleName}

Private bType As B{MockVbeBuilder.TestModuleName}
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.StandardModule, out _);
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName("mTest").First();

                var resolver = Support.SetupResolver(state, rewritingManager, null);
                var modelfactory = resolver.Resolve<IEncapsulateFieldModelFactory>();
                var model = modelfactory.Create(target);

                var field = model["mTest"];
                field.PropertyIdentifier = "ATest";
                model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;

                var test = model.ObjectStateUDTCandidates;
                Assert.AreEqual(3, test.Count());

                var previewProvider = resolver.Resolve<EncapsulateFieldPreviewProvider>();

                var firstPreview = previewProvider.Preview(model);
                StringAssert.Contains("Property Get ATest", firstPreview);

                field.PropertyIdentifier = "BTest";
                var secondPreview = previewProvider.Preview(model);
                StringAssert.Contains("Property Get BTest", secondPreview);
                StringAssert.DoesNotContain("Property Get ATest", secondPreview);
            }
        }

        [TestCase(EncapsulateFieldStrategy.UseBackingFields)]
        [TestCase(EncapsulateFieldStrategy.ConvertFieldsToUDTMembers)]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldPreviewProvider))]
        public void Preview_IncludeEndOfChangesMarker(EncapsulateFieldStrategy strategy)
        {
            var inputCode =
$@"Option Explicit

Public mTest As Long
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.StandardModule, out _);
            (RubberduckParserState state, IRewritingManager rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName("mTest").First();

                var resolver = Support.SetupResolver(state, rewritingManager, null);
                var modelfactory = resolver.Resolve<IEncapsulateFieldModelFactory>();
                var previewProvider = resolver.Resolve<EncapsulateFieldPreviewProvider>();

                var model = modelfactory.Create(target);

                model.EncapsulateFieldStrategy = strategy;

                var previewResult = previewProvider.Preview(model);

                StringAssert.Contains(RefactoringsUI.EncapsulateField_PreviewMarker, previewResult);
            }
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            return Support.SupportTestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }
    }
}
