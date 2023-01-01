using System;
using System.Linq;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodRefactoringAction : CodeOnlyRefactoringActionBase<ExtractMethodModel>
    {

        public ExtractMethodRefactoringAction(IRewritingManager rewritingManager) : base(rewritingManager)
        {

        }

        public override void Refactor(ExtractMethodModel model, IRewriteSession rewriteSession)
        {
            var selection = model.QualifiedSelection;
            var selectedContexts = model.SelectedContexts;

            var rewriter = rewriteSession.CheckOutModuleRewriter(selection.QualifiedName);
            var startIndex = selectedContexts.First().Start.TokenIndex;
            var endIndex = selectedContexts.Last().Stop.TokenIndex;
            var selectionInterval = new Interval(startIndex, endIndex);

            rewriter.InsertAfter(model.TargetMethod.Context.Stop.TokenIndex, Environment.NewLine + model.NewMethodCode);
            rewriter.Replace(selectionInterval, model.ReplacementCode + Environment.NewLine);

            rewriter.Selection = new Selection(selection.Selection.StartLine, selection.Selection.StartColumn);
        }
    }
}