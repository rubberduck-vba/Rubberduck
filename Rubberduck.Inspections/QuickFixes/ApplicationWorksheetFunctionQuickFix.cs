using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ApplicationWorksheetFunctionQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> {typeof(ApplicationWorksheetFunctionInspection) };

        public ApplicationWorksheetFunctionQuickFix(RubberduckParserState state)
        {
            _state = state;
        }
        
        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "WorksheetFunction.");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ApplicationWorksheetFunctionQuickFix;
        }

        public bool CanFixInProcedure { get; } = true;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}
