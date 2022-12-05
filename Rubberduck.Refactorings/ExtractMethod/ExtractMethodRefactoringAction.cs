using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodRefactoringAction : CodeOnlyRefactoringActionBase<ExtractMethodModel>
    {

        public ExtractMethodRefactoringAction(IRewritingManager rewritingManager) : base(rewritingManager)
        {

        }

        public override void Refactor(ExtractMethodModel model, IRewriteSession rewriteSession)
        {
            var selection = model.Selection;

        }
    }
}