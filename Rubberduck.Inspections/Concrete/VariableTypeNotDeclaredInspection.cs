using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class VariableTypeNotDeclaredInspection : InspectionBase
    {
        public VariableTypeNotDeclaredInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where (item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant
                            || (item.DeclarationType == DeclarationType.Parameter && !item.IsArray))
                         && !item.IsTypeSpecified
                         && !item.IsUndeclared
                         select new VariableTypeNotDeclaredInspectionResult(this, item);

            return issues;
        }
    }
}
