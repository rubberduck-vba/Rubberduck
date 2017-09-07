using System;
using System.Collections.Generic;
using System.Linq;
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

            UpdateContext((VBAParser.ElseBlockContext)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            VBAParser.BlockContext elseBlock = context.block();

            if (elseBlock.ChildCount == 0 )
            {
                rewriter.Remove(context);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveEmptyElseBlockQuickFix;
        }

        public bool CanFixInProcedure { get; }  = false;
        public bool CanFixInModule { get; } = false;
        public bool CanFixInProject { get; } = false;
    }
}
