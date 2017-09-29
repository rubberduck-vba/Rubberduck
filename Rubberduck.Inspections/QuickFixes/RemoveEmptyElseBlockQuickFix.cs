using Antlr4.Runtime.Tree;
using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;


namespace Rubberduck.Inspections.QuickFixes
{
    class RemoveEmptyElseBlockQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(EmptyElseBlockInspection) };
        private readonly RubberduckParserState _state;

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public RemoveEmptyElseBlockQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public void Fix(IInspectionResult result)
        {
            IModuleRewriter rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            //dynamic used since it's not known at run-time
            UpdateContext((dynamic)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            VBAParser.BlockContext elseBlock = context.block();

            if (elseBlock.ChildCount == 0 )
            {
                //string rewrittenBlock = AdjustedBlockText(context.block());
                //rewriter.InsertBefore(context.start.StartIndex, rewrittenBlock);
                rewriter.Remove(context);
            }
            /*
             * There isn't any need to invert the condition since its
             * only the else block thats empty. IE: it doesn't affect
             * the TRUE portion that preceeds it.
             */
        }

        //private bool FirstBlockStmntHasLabel(VBAParser.BlockContext block)
        //    => block.blockStmt()?.FirstOrDefault()?.statementLabelDefinition() != null;

        //private bool BlockHasDeclaration(VBAParser.BlockContext block)
        //    => block.blockStmt()?.Any() ?? false;


        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveEmptyElseBlockQuickFix;
        }

        public bool CanFixInProcedure { get; }  = false;
        public bool CanFixInModule { get; } = false;
        public bool CanFixInProject { get; } = false;
    }
}
