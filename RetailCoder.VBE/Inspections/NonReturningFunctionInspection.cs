using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public sealed class NonReturningFunctionInspection : InspectionBase
    {
        public NonReturningFunctionInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return RubberduckUI.NonReturningFunction_; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var declarations = UserDeclarations.ToList();

            var interfaceMembers = declarations.FindInterfaceMembers();
            var interfaceImplementationMembers = declarations.FindInterfaceImplementationMembers();

            var functions = declarations
                .Where(declaration => ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var issues = functions
                .Where(declaration => declaration.References.All(r => !r.IsAssignment))
                .Select(issue => new NonReturningFunctionInspectionResult(this, string.Format(Description, issue.IdentifierName), new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context), interfaceImplementationMembers.Select(m => m.Scope).Contains(issue.Scope)));

            return issues;
        }
    }
}