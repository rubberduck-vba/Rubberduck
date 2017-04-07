using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class MakeSingleLineParameterQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MultilineParameterInspection)
        };

        public MakeSingleLineParameterQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var parameter = result.Context.GetText()
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            rewriter.Replace(result.Target, parameter);
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.MakeSingleLineParameterQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}
