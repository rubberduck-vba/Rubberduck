using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObjectVariableNotSetInspection : InspectionBase
    {
        private readonly VariableRequiresSetAssignmentEvaluator _setRequirementEvaluator;

        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
            _setRequirementEvaluator = new VariableRequiresSetAssignmentEvaluator(state);
        }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var allInterestingDeclarations =
                _setRequirementEvaluator.GetDeclarationsPotentiallyRequiringSetAssignment();

            var candidateReferencesRequiringSetAssignment = 
                allInterestingDeclarations.SelectMany(dec => dec.References);

            var referencesRequiringSetAssignment = candidateReferencesRequiringSetAssignment                  
                .Where(reference => _setRequirementEvaluator.RequiresSetAssignment(reference));

            var objectVariableNotSetReferences = referencesRequiringSetAssignment.Where(reference => FlagIfObjectVariableNotSet(reference));

            return objectVariableNotSetReferences.Select(reference => new ObjectVariableNotSetInspectionResult(this, reference));
        }

        private bool FlagIfObjectVariableNotSet(IdentifierReference reference)
        {
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
            return (reference.IsAssignment && letStmtContext != null && letStmtContext.LET() == null);
        }
    }
}
