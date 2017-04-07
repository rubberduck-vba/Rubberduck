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
    public class ChangeParameterByRefByValQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(ImplicitByRefParameterInspection) };
        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        private readonly RubberduckParserState _state;

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