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
        public ObjectVariableNotSetInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) {  }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var allInterestingDeclarations =
                VariableRequiresSetAssignmentEvaluator.GetDeclarationsPotentiallyRequiringSetAssignment(State.AllUserDeclarations);

            var candidateReferencesRequiringSetAssignment = 
                allInterestingDeclarations.SelectMany(dec => dec.References)
                    .Where(dec => !IsIgnoringInspectionResultFor(dec, AnnotationName))
                    .Where(reference => reference.IsAssignment);

            var referencesRequiringSetAssignment = candidateReferencesRequiringSetAssignment                  
                .Where(reference => VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(reference, State.AllUserDeclarations));

            var objectVariableNotSetReferences = referencesRequiringSetAssignment.Where(reference => FlagIfObjectVariableNotSet(reference));

            return objectVariableNotSetReferences.Select(reference => new ObjectVariableNotSetInspectionResult(this, reference));
        }

        private bool FlagIfObjectVariableNotSet(IdentifierReference reference)
        {
            var letStmtContext = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(reference.Context);
            return (reference.IsAssignment && letStmtContext != null);
        }
    }
}
