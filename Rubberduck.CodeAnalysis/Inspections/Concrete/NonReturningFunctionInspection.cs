using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about 'Function' and 'Property Get' procedures whose return value is not assigned.
    /// </summary>
    /// <why>
    /// Both 'Function' and 'Property Get' accessors should always return something. Omitting the return assignment is likely a bug.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     Dim foo As Long
    ///     foo = 42
    ///     'function will always return 0
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Function GetFoo() As Long
    ///     Dim foo As Long
    ///     foo = 42
    ///     GetFoo = foo
    /// End Function
    /// ]]>
    /// </example>
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
            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers().ToHashSet();

            var functions = State.DeclarationFinder.UserDeclarations(DeclarationType.Function)
                .Where(declaration => !interfaceMembers.Contains(declaration));

            var unassigned = functions.Where(function => IsReturningUserDefinedType(function) 
                                                            && !IsUserDefinedTypeAssigned(function)
                                                          || !IsReturningUserDefinedType(function) 
                                                            && !IsAssigned(function));

            return unassigned
                .Select(issue =>
                    new DeclarationInspectionResult(this,
                                         string.Format(InspectionResults.NonReturningFunctionInspection, issue.IdentifierName),
                                         issue));
        }

        private bool IsAssigned(Declaration function)
        {
            var inScopeIdentifierReferences = function.References.Where(r => r.ParentScoping.Equals(function));
            return inScopeIdentifierReferences.Any(reference => reference.IsAssignment 
                                                                || IsAssignedByRefArgument(function, reference));
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

        private static bool IsReturningUserDefinedType(Declaration member)
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
