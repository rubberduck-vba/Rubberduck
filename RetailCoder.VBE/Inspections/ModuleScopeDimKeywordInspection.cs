using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ModuleScopeDimKeywordInspection : InspectionBase
    {
        public ModuleScopeDimKeywordInspection(RubberduckParserState state) 
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ModuleScopeDimKeywordInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ModuleScopeDimKeywordInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        private static readonly IReadOnlyList<DeclarationType> ModuleTypes = new[]
        {
            DeclarationType.ProceduralModule, 
            DeclarationType.ClassModule, 
            DeclarationType.UserForm, 
            DeclarationType.Document, 
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
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