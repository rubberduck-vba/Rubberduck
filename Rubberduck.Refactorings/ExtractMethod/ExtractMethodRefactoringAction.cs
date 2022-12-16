using System.Linq;
using Antlr4.Runtime.Misc;
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
            var selectedContexts = model.SelectedContexts;

            var rewriter = rewriteSession.CheckOutModuleRewriter(selection.QualifiedName);
            var startIndex = selectedContexts.First().Start.TokenIndex;
            var endIndex = selectedContexts.Last().Stop.TokenIndex;
            var selectionInterval = new Interval(startIndex, endIndex);

            rewriter.InsertAfter(model.TargetMethod.Context.Stop.TokenIndex, model.PreviewCode);
            rewriter.Replace(selectionInterval, model.ReplacementCode);
        }
    }
}