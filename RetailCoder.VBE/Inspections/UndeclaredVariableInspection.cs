using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class UndeclaredVariableInspection : InspectionBase
    {
        public UndeclaredVariableInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.UndeclaredVariableInspectionMeta; } }
        public override string Description { get { return InspectionsUI.UndeclaredVariableInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return UserDeclarations.Where(item => item.IsUndeclared && item.DeclarationType == DeclarationType.Variable)
                .Select(item => new UndeclaredVariableInspectionResult(this, item));
        }
    }
}