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
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables that are never assigned.
    /// </summary>
    /// <why>
    /// A variable that is never assigned is probably a sign of a bug. 
    /// This inspection may yield false positives if the variable is assigned through a ByRef parameter assignment, or 
    /// if UserForm controls fail to resolve, references to these controls in code-behind can be flagged as unassigned and undeclared variables.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long ' declared, but not assigned
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
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
                    && !declaration.References.Any(reference => reference.IsAssignment || IsAssignedByRefArgument(reference.ParentScoping, reference)));

            return declarations.Select(issue => 
                new DeclarationInspectionResult(this, string.Format(InspectionResults.VariableNotAssignedInspection, issue.IdentifierName), issue));
        }

        private bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference)
        {
            var argExpression = reference.Context.GetAncestor<VBAParser.ArgumentExpressionContext>();
            var parameter = State.DeclarationFinder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argExpression, enclosingProcedure);

            // note: not recursive, by design.
            return parameter != null
                   && (parameter.IsImplicitByRef || parameter.IsByRef)
                   && parameter.References.Any(r => r.IsAssignment);
        }
    }
}
