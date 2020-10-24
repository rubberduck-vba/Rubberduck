using System;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public abstract class EncapsulateFieldInteractiveRefactoringTest : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        //RefactoringTestBase.NoActiveSelection_Throws passes a null
        //IDeclarationFinderProvider parameter to 'TestRefactoring(...).  
        //The EncapsulateFieldRefactoring tests Resolver throws a different
        //exception type without a valid interface reference and causes the 
        //base class version of the test to fail.
        [Test]
        [Category("Refactorings")]
        public override void NoActiveSelection_Throws()
        {
            var testVbe = TestVbe(string.Empty, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(testVbe);
            using (state)
            {
                var refactoring = TestRefactoring(rewritingManager, state, initialSelection: null);
                Assert.Throws<NoActiveSelectionException>(() => refactoring.Refactor());
            }
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction, ISelectionService selectionService)
        {
            throw new NotImplementedException();
        }
    }
}
