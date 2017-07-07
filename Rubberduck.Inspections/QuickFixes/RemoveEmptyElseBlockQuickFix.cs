using Antlr4.Runtime.Tree;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
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
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            ITerminalNode elseBlock = context.ELSE();

            if (BlockHasDeclaration(context.block()))
            {
                string rewrittenBlock = AdjustedBlockText(context.block());
                rewriter.InsertBefore(context.start.StartIndex, rewrittenBlock);
            }



        }

        private string AdjustedBlockText(VBAParser.BlockContext blockContext)
        {
            string blockText = blockContext.GetText();
            
            if (FirstBlockStmntHasLabel(blockContext))
            {
                blockText = Environment.NewLine + blockContext;
            }

            return blockText;
        }

        private bool FirstBlockStmntHasLabel(VBAParser.BlockContext block)
            => block.blockStmt()?.FirstOrDefault()?.statementLabelDefinition() != null;

        private bool BlockHasDeclaration(VBAParser.BlockContext block)
            => block.blockStmt()?.Any() ?? false;


        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveEmptyElseBlockQuickFix;
        }

        public bool CanFixInProcedure { get; }  = false;
        public bool CanFixInModule { get; } = false;
        public bool CanFixInProject { get; } = false;
    }
}
