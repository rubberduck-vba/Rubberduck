using System.Linq;
using Antlr4.Runtime.Tree;
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
    /// Warns about 'Function' and 'Property Get' procedures whose return value is not assigned.
    /// </summary>
    /// <why>
    /// Both 'Function' and 'Property Get' accessors should always return something. Omitting the return assignment is likely a bug.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     Dim foo As Long
    ///     foo = 42
    ///     'function will always return 0
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     Dim foo As Long
    ///     foo = 42
    ///     GetFoo = foo
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class NonReturningFunctionInspection : DeclarationInspectionBase
    {
        public NonReturningFunctionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Function)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration is ModuleBodyElementDeclaration member
                   && !member.IsInterfaceMember
                   && (IsReturningUserDefinedType(member)
                       && !IsUserDefinedTypeAssigned(member)
                       || !IsReturningUserDefinedType(member)
                       && !IsAssigned(member, finder));
        }

        private bool IsAssigned(Declaration member, DeclarationFinder finder)
        {
            var inScopeIdentifierReferences = member.References
                .Where(reference => reference.ParentScoping.Equals(member));
            return inScopeIdentifierReferences
                .Any(reference => reference.IsAssignment 
                                  || IsAssignedByRefArgument(member, reference, finder));
        }

        private bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference, DeclarationFinder finder)
        {
            var argExpression = ImmediateArgumentExpressionContext(reference);

            if (argExpression is null)
            {
                return false;
            }

            var argument = argExpression.GetAncestor<VBAParser.ArgumentContext>();
            var parameter = finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(argument, enclosingProcedure);

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

        private static bool IsReturningUserDefinedType(Declaration member)
        {
            return member.AsTypeDeclaration != null 
                   && member.AsTypeDeclaration.DeclarationType == DeclarationType.UserDefinedType;
        }

        private static bool IsUserDefinedTypeAssigned(Declaration member)
        {
            // ref. #2257:
            // A function returning a UDT type shouldn't trip this inspection if
            // at least one UDT member is assigned a value.
            var block = member.Context.GetChild<VBAParser.BlockContext>(0);
            var visitor = new FunctionReturnValueAssignmentLocator(member.IdentifierName);
            var result = visitor.VisitBlock(block);
            return result;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.NonReturningFunctionInspection, declaration.IdentifierName);
        }

        /// <summary>
        /// A visitor that visits a member's body and returns <c>true</c> if any <c>LET</c> statement (assignment) is assigning the specified <c>name</c>.
        /// </summary>
        private class FunctionReturnValueAssignmentLocator : VBAParserBaseVisitor<bool>
        {
            private readonly string _name;
            private bool _inFunctionReturnWithExpression;

            public FunctionReturnValueAssignmentLocator(string name)
            {
                _name = name;
                _inFunctionReturnWithExpression = false;
            }

            protected override bool DefaultResult => false;

            protected override bool ShouldVisitNextChild(IRuleNode node, bool currentResult)
            {
                return !currentResult;
            }

            //This is actually the default implementation, but for explicities sake stated here.
            protected override bool AggregateResult(bool aggregate, bool nextResult)
            {
                return nextResult;
            }

            public override bool VisitWithStmt(VBAParser.WithStmtContext context)
            {
                var oldInFunctionReturnWithExpression = _inFunctionReturnWithExpression;
                _inFunctionReturnWithExpression = context.expression().GetText() == _name;
                var result = base.VisitWithStmt(context);
                _inFunctionReturnWithExpression = oldInFunctionReturnWithExpression;
                return result;
            }

            public override bool VisitLetStmt(VBAParser.LetStmtContext context)
            {
                var LHS = context.lExpression();
                if (_inFunctionReturnWithExpression
                        && LHS is VBAParser.WithMemberAccessExprContext)
                {
                    return true;
                }
                var leftmost = LHS.GetChild(0).GetText();
                return leftmost == _name;
            }

            public override bool VisitSetStmt(VBAParser.SetStmtContext context)
            {
                var LHS = context.lExpression();
                if (_inFunctionReturnWithExpression
                        && LHS is VBAParser.WithMemberAccessExprContext)
                {
                    return true;
                }
                var leftmost = LHS.GetChild(0).GetText();
                return leftmost == _name;
            }
        }
    }
}
