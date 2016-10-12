using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class NonReturningFunctionInspection : InspectionBase
    {
        public NonReturningFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.NonReturningFunctionInspectionMeta; }}
        public override string Description { get { return InspectionsUI.NonReturningFunctionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var interfaceMembers = declarations.FindInterfaceMembers();
            var interfaceImplementationMembers = declarations.FindInterfaceImplementationMembers();

            var functions = declarations
                .Where(declaration => ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var unassigned = from function in functions
                             let isUdt = IsReturningUserDefinedType(function)
                             let inScopeRefs = function.References.Where(r => r.ParentScoping.Equals(function))
                             where (!isUdt && (!inScopeRefs.Any(r => r.IsAssignment)))
                                || (isUdt && !IsUserDefinedTypeAssigned(function))
                             select function;

            return unassigned
                .Select(issue =>
                    new NonReturningFunctionInspectionResult(this, new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context),
                            interfaceImplementationMembers.Select(m => m.Scope).Contains(issue.Scope), issue));
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
                _result = leftmost == _name;
                return _result;
            }
        }
    }
}
