﻿using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class VariableTypeNotDeclaredInspection : InspectionBase
    {
        public VariableTypeNotDeclaredInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.VariableTypeNotDeclaredInspectionMeta; } }
        public override string Description { get { return InspectionsUI.VariableTypeNotDeclaredInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = from item in UserDeclarations
                         where (item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant
                            || (item.DeclarationType == DeclarationType.Parameter && !item.IsArray))
                         && !item.IsTypeSpecified
                         select new VariableTypeNotDeclaredInspectionResult(this, item);

            return issues;
        }
    }
}
