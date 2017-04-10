using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public sealed class RemoveUnusedDeclarationQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ConstantNotUsedInspection),
            typeof(ProcedureNotUsedInspection),
            typeof(VariableNotUsedInspection)
        };

        public RemoveUnusedDeclarationQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.Remove(result.Target);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveUnusedDeclarationQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}