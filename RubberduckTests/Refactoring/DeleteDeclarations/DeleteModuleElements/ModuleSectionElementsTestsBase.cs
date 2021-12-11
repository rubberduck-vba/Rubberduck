using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    public class ModuleSectionElementsTestsBase
    {
        protected readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();
        protected string GetRetainedCodeBlock(string moduleCode, Func<RubberduckParserState, IEnumerable<Declaration>> targetListBuilder, Action<IDeleteDeclarationsModel> modelFlagAction = null)
        {
            var refactoredCode = _support.TestRefactoring(
                targetListBuilder,
                RefactorModuleElements,
                modelFlagAction ?? _support.DefaultModelFlagAction,
                (MockVbeBuilder.TestModuleName, moduleCode, ComponentType.StandardModule));

            return refactoredCode[MockVbeBuilder.TestModuleName];
        }

        protected static IExecutableRewriteSession RefactorModuleElements(RubberduckParserState state, IEnumerable<Declaration> targets, IRewritingManager rewritingManager, Action<IDeleteDeclarationsModel> modelFlagAction)
        {
            var model = new DeleteModuleElementsModel(targets);
            modelFlagAction(model);

            var session = rewritingManager.CheckOutCodePaneSession();

            var refactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteModuleElementsRefactoringAction>();

            refactoringAction.Refactor(model, session);

            return session;
        }
    }
}
