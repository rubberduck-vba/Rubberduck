using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UnassignedVariableUsageInspection : InspectionBase
    {
        public UnassignedVariableUsageInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var declarations = UserDeclarations.Where(declaration => 
                declaration.DeclarationType == DeclarationType.Variable
                && !UserDeclarations.Any(d => d.DeclarationType == DeclarationType.UserDefinedType
                    && d.IdentifierName == declaration.AsTypeName)
                && !declaration.IsSelfAssigned
                && !declaration.References.Any(reference => reference.IsAssignment && !IsIgnoringInspectionResultFor(reference, AnnotationName)));

            //The parameter scoping was apparently incorrect before - need to filter for the actual function.
            var lenFunction = BuiltInDeclarations.SingleOrDefault(s => s.DeclarationType == DeclarationType.Function && s.Scope.Equals("VBE7.DLL;VBA.Strings.Len"));
            var lenbFunction = BuiltInDeclarations.SingleOrDefault(s => s.DeclarationType == DeclarationType.Function && s.Scope.Equals("VBE7.DLL;VBA.Strings.Len"));

            return declarations.Where(d => d.References.Any() &&
                                           !DeclarationReferencesContainsReference(lenFunction, d) &&
                                           !DeclarationReferencesContainsReference(lenbFunction, d))
                               .SelectMany(d => d.References)
                               .Select(r => new UnassignedVariableUsageInspectionResult(this, r));
        }

        private bool DeclarationReferencesContainsReference(Declaration parentDeclaration, Declaration target)
        {
            if (parentDeclaration == null)
            {
                return false;
            }
            
            foreach (var targetReference in target.References)
            {
                foreach (var reference in parentDeclaration.References)
                {
                    var context = (ParserRuleContext) reference.Context.Parent;
                    if (context.GetSelection().Contains(targetReference.Selection))
                    {
                        return true;
                    }
                }
            }
            
            return false;
        }
    }
}
