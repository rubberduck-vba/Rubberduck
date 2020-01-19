using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns when a user function's return value is discarded at all its call sites.
    /// </summary>
    /// <why>
    /// A 'Function' procedure normally means its return value to be captured and consumed by the calling code. 
    /// It's possible that not all call sites need the return value, but if the value is systematically discarded then this
    /// means the function is side-effecting, and thus should probably be a 'Sub' procedure instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     GetFoo ' return value is not captured
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     GetFoo ' return value is discarded
    /// End Sub
    /// 
    /// Public Sub DoSomethingElse()
    ///     Dim foo As Long
    ///     foo = GetFoo ' return value is captured
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class FunctionReturnValueAlwaysDiscardedInspection : DeclarationInspectionBase
    {
        public FunctionReturnValueAlwaysDiscardedInspection(RubberduckParserState state)
            :base(state, DeclarationType.Function)
        {}

        protected override bool IsResultDeclaration(Declaration declaration)
        {
            if (!(declaration is ModuleBodyElementDeclaration moduleBodyElementDeclaration))
            {
                return false;
            }

            //We only report the interface itself.
            if (moduleBodyElementDeclaration.IsInterfaceImplementation)
            {
                return false;
            }

            var finder = DeclarationFinderProvider.DeclarationFinder;

            if (moduleBodyElementDeclaration.IsInterfaceMember)
            {
                return IsInterfaceIssue(moduleBodyElementDeclaration, finder);
            }

            return IsIssueItself(moduleBodyElementDeclaration);
        }

        private bool IsIssueItself(ModuleBodyElementDeclaration declaration)
        {
            var procedureCallReferences = ProcedureCallReferences(declaration).ToHashSet();
            if (!procedureCallReferences.Any())
            {
                return false;
            }

            return declaration.References
                .All(reference => procedureCallReferences.Contains(reference)
                                  || reference.IsAssignment && IsReturnStatement(declaration, reference));
        }

        private bool IsReturnStatement(Declaration function, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(function) && assignment.Declaration.Equals(function);
        }

        private bool IsInterfaceIssue(ModuleBodyElementDeclaration declaration, DeclarationFinder finder)
        {
            if (!IsIssueItself(declaration))
            {
                return false;
            }

            var implementations = finder.FindInterfaceImplementationMembers(declaration);
            return implementations.All(implementation => IsIssueItself(implementation)
                                                         || implementation.References.All(reference =>
                                                             reference.IsAssignment
                                                             && IsReturnStatement(implementation, reference)));
        }

        private static IEnumerable<IdentifierReference> ProcedureCallReferences(Declaration declaration)
        {
            return declaration.References
                .Where(IsProcedureCallReference);
        }

        private static bool IsProcedureCallReference(IdentifierReference reference)
        {
            return reference?.Declaration != null
                   && !reference.IsAssignment
                   && !reference.IsArrayAccess
                   && !reference.IsInnerRecursiveDefaultMemberAccess
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

        protected override string ResultDescription(Declaration declaration)
        {
            var functionName = declaration.QualifiedName.ToString();
            return string.Format(InspectionResults.FunctionReturnValueAlwaysDiscardedInspection, functionName);
        }
    }
}
