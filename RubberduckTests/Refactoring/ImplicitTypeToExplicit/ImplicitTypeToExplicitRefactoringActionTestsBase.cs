using System;
using System.Linq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactoring.ParseTreeValue;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;

namespace RubberduckTests.Refactoring.ImplicitTypeToExplicit
{
    public class ImplicitTypeToExplicitRefactoringActionTestsBase : RefactoringActionTestBase<ImplicitTypeToExplicitModel>
    {
        protected override IRefactoringAction<ImplicitTypeToExplicitModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            return new ImplicitTypeToExplicitRefactoringAction(state, new ParseTreeValueFactory(), rewritingManager);
        }

        protected string RefactoredCode(string inputCode, (string targetName, DeclarationType declarationType) tuple)
        {
            return RefactoredCode(inputCode,
                state => TestModel(state, tuple, (model) => model));
        }

        protected static ImplicitTypeToExplicitModel TestModel(IDeclarationFinderProvider state, (string targetIdentifier, DeclarationType declarationType) tuple, Func<ImplicitTypeToExplicitModel, ImplicitTypeToExplicitModel> modelAdjustment)
        {
            var target = state.DeclarationFinder.UserDeclarations(tuple.declarationType)
                .Single(d => d.IdentifierName == tuple.targetIdentifier);
            var model = new ImplicitTypeToExplicitModel(target);
            return modelAdjustment(model);
        }
    }
}
