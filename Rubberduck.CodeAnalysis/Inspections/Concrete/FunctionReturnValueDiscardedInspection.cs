using Antlr4.Runtime;
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
    /// Warns when a user function's return value is not used at a call site.
    /// </summary>
    /// <why>
    /// A 'Function' procedure normally means its return value to be captured and consumed by the calling code. 
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     GetFoo ' return value is not captured
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = GetFoo
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class FunctionReturnValueDiscardedInspection : IdentifierReferenceInspectionBase
    {
        public FunctionReturnValueDiscardedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference?.Declaration != null
                   && reference.Declaration.IsUserDefined
                   && !reference.IsAssignment
                   && !reference.IsArrayAccess
                   && !reference.IsInnerRecursiveDefaultMemberAccess
                   && reference.Declaration.DeclarationType == DeclarationType.Function
                   && IsCalledAsProcedure(reference.Context);
        }

        private static bool IsCalledAsProcedure(ParserRuleContext context)
        {
            var callStmt = context.GetAncestor<VBAParser.CallStmtContext>();
            if (callStmt == null)
            {
                return false;
            }

            //If we are in an argument list, the value is used somewhere in defining the argument.
            var argumentListParent = context.GetAncestor<VBAParser.ArgumentListContext>();
            if (argumentListParent != null)
            {
                return false;
            }

            //Member accesses are parsed right-to-left, e.g. 'foo.Bar' is the parent of 'foo'.
            //Thus, having a member access parent means that the return value is used somehow.
            var ownFunctionCallExpression = context.Parent is VBAParser.MemberAccessExprContext methodCall
                ? methodCall
                : context;
            var memberAccessParent = ownFunctionCallExpression.GetAncestor<VBAParser.MemberAccessExprContext>();
            if (memberAccessParent != null)
            {
                return false;
            }

            //If we are in an output list, the value is used somewhere in defining the argument.
            var outputListParent = context.GetAncestor<VBAParser.OutputListContext>();
            return outputListParent == null;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var functionName = reference.Declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.FunctionReturnValueDiscardedInspection, functionName);
        }
    }
}
