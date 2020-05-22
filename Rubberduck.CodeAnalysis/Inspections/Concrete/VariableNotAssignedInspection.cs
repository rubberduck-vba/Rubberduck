using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about variables that are never assigned.
    /// </summary>
    /// <why>
    /// A variable that is never assigned is probably a sign of a bug. 
    /// This inspection may yield false positives if the variable is assigned through a ByRef parameter assignment, or 
    /// if UserForm controls fail to resolve, references to these controls in code-behind can be flagged as unassigned and undeclared variables.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long ' declared, but not assigned
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     value = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class VariableNotAssignedInspection : DeclarationInspectionBase
    {
        public VariableNotAssignedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Variable)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration != null
                   && !declaration.IsArray // ignore arrays. todo: ArrayIndicesNotAccessedInspection
                   && !declaration.IsWithEvents
                   && !declaration.IsSelfAssigned
                   && !HasUdtType(declaration, finder) // UDT variables don't need to be assigned
                   && !declaration.References.Any(reference => reference.IsAssignment || IsAssignedByRefArgument(reference.ParentScoping, reference, finder));
        }

        private static bool HasUdtType(Declaration declaration, DeclarationFinder finder)
        {
            return finder.MatchName(declaration.AsTypeName)
                .Any(item => item.DeclarationType == DeclarationType.UserDefinedType);
        }

        private static bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference, DeclarationFinder finder)
        {
            var argExpression = ImmediateArgumentExpressionContext(reference);

            if (argExpression is null)
            {
                return false;
            }

            var parameter = finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argExpression, enclosingProcedure);

            // note: not recursive, by design.
            return parameter != null
                   && (parameter.IsImplicitByRef || parameter.IsByRef)
                   && parameter.References.Any(r => r.IsAssignment);
        }

        private static VBAParser.ArgumentExpressionContext ImmediateArgumentExpressionContext(IdentifierReference reference)
        {
            var context = reference.Context;
            //The context is either already a simpleNameExprContext or an IdentifierValueContext used in a sub-rule of some other lExpression alternative. 
            var lExpressionNameContext = context is VBAParser.SimpleNameExprContext simpleName
                ? simpleName
                : context.GetAncestor<VBAParser.LExpressionContext>();

            //To be an immediate argument and, thus, assignable by ref, the structure must be argumentExpression -> expression -> lExpression.
            return lExpressionNameContext?
                .Parent?
                .Parent as VBAParser.ArgumentExpressionContext;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.VariableNotAssignedInspection, declaration.IdentifierName);
        }
    }
}
