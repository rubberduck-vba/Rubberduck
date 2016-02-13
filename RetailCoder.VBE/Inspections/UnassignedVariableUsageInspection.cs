using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class UnassignedVariableUsageInspection : InspectionBase
    {
        public UnassignedVariableUsageInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public override string Description { get { return RubberduckUI.UnassignedVariableUsage_; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var usages = UserDeclarations.Where(declaration => 
                declaration.DeclarationType == DeclarationType.Variable
                && !UserDeclarations.Any(d => d.DeclarationType == DeclarationType.UserDefinedType
                    && d.IdentifierName == declaration.AsTypeName)
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment))
                .SelectMany(declaration => declaration.References)
                .Where(usage => !usage.IsInspectionDisabled(AnnotationName));

            foreach (var issue in usages)
            {
                yield return new UnassignedVariableUsageInspectionResult(this, string.Format(Description, issue.Context.GetText()), issue.Context, issue.QualifiedModuleName);
            }
        }
    }
}