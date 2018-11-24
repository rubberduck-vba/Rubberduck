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
    public sealed class NonReturningFunctionInspection : InspectionBase
    {
        public NonReturningFunctionInspection(RubberduckParserState state)
            : base(state) { }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers();

            var functions = declarations
                .Where(declaration => ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var unassigned = (from function in functions
                             let isUdt = IsReturningUserDefinedType(function)
                             let inScopeRefs = function.References.Where(r => r.ParentScoping.Equals(function))
                             where (!isUdt && (!inScopeRefs.Any(r => r.IsAssignment) && 
                                               !inScopeRefs.Any(reference => IsAssignedByRefArgument(function, reference))))
                                || (isUdt && !IsUserDefinedTypeAssigned(function))
                             select function)
                             .ToList();

            return unassigned
                .Select(issue =>
                    new DeclarationInspectionResult(this,
                                         string.Format(InspectionResults.NonReturningFunctionInspection, issue.IdentifierName),
                                         issue));
        }

        private bool IsAssignedByRefArgument(Declaration enclosingProcedure, IdentifierReference reference)
        {
            var argExpression = reference.Context.GetAncestor<VBAParser.ArgumentExpressionContext>();
            if (argExpression?.GetDescendent<VBAParser.ParenthesizedExprContext>() != null || argExpression?.BYVAL() != null)
            {
                // not an argument, or argument is parenthesized and thus passed ByVal
                return false;
            }

            var callStmt = argExpression?.GetAncestor<VBAParser.CallStmtContext>();
            var procedureName = callStmt?.GetDescendent<VBAParser.LExpressionContext>()
                                         .GetDescendents<VBAParser.IdentifierContext>()
                                         .LastOrDefault()?.GetText();
            if (procedureName == null)
            {
                // if we don't know what we're calling, we can't dig any further
                return false;
            }

            var procedure = State.DeclarationFinder.MatchName(procedureName)
                .Where(p => AccessibilityCheck.IsAccessible(enclosingProcedure, p))
                .SingleOrDefault(p => !p.DeclarationType.HasFlag(DeclarationType.Property) || p.DeclarationType.HasFlag(DeclarationType.PropertyGet));
            var parameters = State.DeclarationFinder.Parameters(procedure);

            ParameterDeclaration parameter;
            var namedArg = argExpression.GetAncestor<VBAParser.NamedArgumentContext>();
            if (namedArg != null)
            {
                // argument is named: we're lucky
                var parameterName = namedArg.unrestrictedIdentifier().GetText();
                parameter = parameters.SingleOrDefault(p => p.IdentifierName == parameterName);
            }
            else
            {
                // argument is positional: work out its index
                var argList = callStmt.GetDescendent<VBAParser.ArgumentListContext>();
                var args = argList.GetDescendents<VBAParser.PositionalArgumentContext>().ToArray();
                var parameterIndex = args.Select((a, i) =>
                        a.GetDescendent<VBAParser.ArgumentExpressionContext>() == argExpression ? (a, i) : (null, -1))
                    .SingleOrDefault(item => item.a != null).i;
                parameter = parameters.OrderBy(p => p.Selection).Select((p, i) => (p, i))
                    .SingleOrDefault(item => item.i == parameterIndex).p;
            }

            if (parameter == null)
            {
                // couldn't locate parameter
                return false;
            }

            // note: not recursive, by design.
            return (parameter.IsImplicitByRef || parameter.IsByRef)
                && parameter.References.Any(r => r.IsAssignment);
        }

        private bool IsReturningUserDefinedType(Declaration member)
        {
            return member.AsTypeDeclaration != null &&
                   member.AsTypeDeclaration.DeclarationType == DeclarationType.UserDefinedType;
        }

        private bool IsUserDefinedTypeAssigned(Declaration member)
        {
            // ref. #2257:
            // A function returning a UDT type shouldn't trip this inspection if
            // at least one UDT member is assigned a value.
            var block = member.Context.GetChild<VBAParser.BlockContext>(0);
            var visitor = new FunctionReturnValueAssignmentLocator(member.IdentifierName);
            var result = visitor.VisitBlock(block);
            return result;
        }
        
        /// <summary>
        /// A visitor that visits a member's body and returns <c>true</c> if any <c>LET</c> statement (assignment) is assigning the specified <c>name</c>.
        /// </summary>
        private class FunctionReturnValueAssignmentLocator : VBAParserBaseVisitor<bool>
        {
            private readonly string _name;
            private bool _result;

            public FunctionReturnValueAssignmentLocator(string name)
            {
                _name = name;
            }

            public override bool VisitBlock(VBAParser.BlockContext context)
            {
                base.VisitBlock(context);
                return _result;
            }

            public override bool VisitLetStmt(VBAParser.LetStmtContext context)
            {
                var leftmost = context.lExpression().GetChild(0).GetText();
                _result = _result || leftmost == _name;
                return _result;
            }

            public override bool VisitSetStmt(VBAParser.SetStmtContext context)
            {
                var leftmost = context.lExpression().GetChild(0).GetText();
                _result = _result || leftmost == _name;
                return _result;
            }
        }
    }
}
