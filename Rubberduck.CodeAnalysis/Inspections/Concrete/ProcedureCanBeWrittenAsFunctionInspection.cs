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
    /// Warns about 'Sub' procedures that could be refactored into a 'Function'.
    /// </summary>
    /// <why>
    /// Idiomatic VB code uses 'Function' procedures to return a single value. If the procedure isn't side-effecting, consider writing it as a
    /// 'Function' rather than a 'Sub' that returns a result through a 'ByRef' parameter.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef result As Long)
    ///     ' ...
    ///     result = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Function DoSomething() As Long
    ///     ' ...
    ///     DoSomething = 42
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Function DoSomething(ByVal arg As Long) As Long
    ///     ' ...
    ///     DoSomething = 42
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ProcedureCanBeWrittenAsFunctionInspection : DeclarationInspectionBase
    {
        public ProcedureCanBeWrittenAsFunctionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, new []{DeclarationType.Procedure}, new []{DeclarationType.LibraryProcedure, DeclarationType.PropertyLet, DeclarationType.PropertySet})
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ModuleBodyElementDeclaration member)
                || member.IsInterfaceImplementation
                || member.IsInterfaceMember
                || finder.FindEventHandlers().Contains(member)
                || member.Parameters.Count(param => param.IsByRef && !param.IsParamArray) != 1)
            {
                return false;
            }

            var parameter = member.Parameters.First(param => param.IsByRef && !param.IsParamArray);
            var parameterReferences = parameter.References.ToList();

            return parameterReferences.Any(reference => IsAssignment(reference, finder));
        }

        private static bool IsAssignment(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference.IsAssignment
                || IsUsageAsAssignedToByRefArgument(reference, finder);
        }

        private static bool IsUsageAsAssignedToByRefArgument(IdentifierReference reference, DeclarationFinder finder)
        {
            var argExpression = ImmediateArgumentExpressionContext(reference);

            if (argExpression == null)
            {
                return false;
            }

            var argument = argExpression.GetAncestor<VBAParser.ArgumentContext>();
            var parameter = finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, reference.QualifiedModuleName);

            if (parameter == null)
            {
                //We have no idea what parameter it is passed to as argument. So, we have to err on the safe side and assume it is not passed by reference.
                return false;
            }

            //We go only one level deep and make a conservative check to avoid a potentially costly recursion.
            return parameter.IsByRef
                && parameter.References.Any(paramReference => paramReference.IsAssignment);
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
            return string.Format(
                InspectionResults.ProcedureCanBeWrittenAsFunctionInspection,
                declaration.IdentifierName);
        }
    }
}
