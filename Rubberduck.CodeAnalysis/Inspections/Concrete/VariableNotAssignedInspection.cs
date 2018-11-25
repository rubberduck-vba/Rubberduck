using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class VariableNotAssignedInspection : InspectionBase
    {
        public VariableNotAssignedInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            // ignore arrays. todo: ArrayIndicesNotAccessedInspection
            var arrays = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable).Where(declaration => declaration.IsArray);

            var declarations = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Except(arrays)
                .Where(declaration =>
                    !declaration.IsWithEvents
                    && State.DeclarationFinder.MatchName(declaration.AsTypeName).All(item => item.DeclarationType != DeclarationType.UserDefinedType) // UDT variables don't need to be assigned
                    && !declaration.IsSelfAssigned
                    && !declaration.References.Any(reference => reference.IsAssignment || IsAssignedByRefArgument(reference.ParentScoping, reference)))
                .Where(result => !IsIgnoringInspectionResultFor(result, AnnotationName));

            return declarations.Select(issue => 
                new DeclarationInspectionResult(this, string.Format(InspectionResults.VariableNotAssignedInspection, issue.IdentifierName), issue));
        }

        private bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference)
        {
            var argExpression = reference.Context.GetAncestor<VBAParser.ArgumentExpressionContext>();
            var parameter = State.DeclarationFinder.FindParameterFromArgument(argExpression, enclosingProcedure);

            // note: not recursive, by design.
            return parameter != null
                   && (parameter.IsImplicitByRef || parameter.IsByRef)
                   && parameter.References.Any(r => r.IsAssignment);
        }
    }
}
