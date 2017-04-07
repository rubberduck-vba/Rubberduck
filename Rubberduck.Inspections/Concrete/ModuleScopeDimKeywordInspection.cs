using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ModuleScopeDimKeywordInspection : InspectionBase
    {
        public ModuleScopeDimKeywordInspection(RubberduckParserState state) 
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[]
        {
            DeclarationType.ProceduralModule, 
            DeclarationType.ClassModule, 
            DeclarationType.UserForm, 
            DeclarationType.Document, 
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var moduleVariables = State.AllUserDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.Variable
                                   && ModuleTypes.Contains(declaration.ParentDeclaration.DeclarationType)
                                   && declaration.Context.Parent.Parent is VBAParser.VariableStmtContext
                                   && ((VBAParser.VariableStmtContext)declaration.Context.Parent.Parent).DIM() != null);
            return moduleVariables.Select(variable => new ModuleScopeDimKeywordInspectionResult(this, variable));
        }
    }
}