using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class UnassignedVariableUsageInspection : InspectionBase
    {
        public UnassignedVariableUsageInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.UnassignedVariableUsageInspectionMeta; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public override string Description { get { return InspectionsUI.UnassignedVariableUsageInspectionName; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var usages = UserDeclarations.Where(declaration => 
                declaration.DeclarationType == DeclarationType.Variable
                && !UserDeclarations.Any(d => d.DeclarationType == DeclarationType.UserDefinedType
                    && d.IdentifierName == declaration.AsTypeName)
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment))
                .SelectMany(declaration => declaration.References)
                .Where(usage => !usage.IsInspectionDisabled(AnnotationName));

            var lenFunction = BuiltInDeclarations.SingleOrDefault(s => s.Scope == "VBE7.DLL;VBA.Strings.Len");
            var lenbFunction = BuiltInDeclarations.SingleOrDefault(s => s.Scope == "VBE7.DLL;VBA.Strings.LenB");

            foreach (var issue in usages)
            {
                if (DeclarationReferencesContainsReference(lenFunction, issue) ||
                    DeclarationReferencesContainsReference(lenbFunction, issue))
                {
                    continue;
                }

                yield return
                    new UnassignedVariableUsageInspectionResult(this, issue.Context, issue.QualifiedModuleName,
                        issue.Declaration);
            }
        }

        private bool DeclarationReferencesContainsReference(Declaration parentDeclaration, IdentifierReference issue)
        {
            if (parentDeclaration == null)
            {
                return false;
            }

            var lenUsesIssue = false;
            foreach (var reference in parentDeclaration.References)
            {
                var context = (ParserRuleContext) reference.Context.Parent;
                if (context.GetSelection().Contains(issue.Selection))
                {
                    lenUsesIssue = true;
                    break;
                }
            }

            if (lenUsesIssue)
            {
                return true;
            }
            return false;
        }
    }
}
