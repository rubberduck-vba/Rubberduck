using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ChangeParameterByRefByValQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(ImplicitByRefParameterInspection) };

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public ChangeParameterByRefByValQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var parameterContext = (VBAParser.ArgContext) result.Target.Context;
            rewriter.InsertBefore(parameterContext.unrestrictedIdentifier().Start.TokenIndex, "ByRef ");
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ImplicitByRefParameterQuickFix;
        }

        public bool CanFixInProcedure { get; } = true;
        public bool CanFixInModule { get; } = true;
        public bool CanFixInProject { get; } = true;
    }
}