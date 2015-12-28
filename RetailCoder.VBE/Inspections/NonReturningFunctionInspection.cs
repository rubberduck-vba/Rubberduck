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
    public class NonReturningFunctionInspection : IInspection
    {
        public NonReturningFunctionInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "NonReturningFunctionInspection"; } }
        public string Description { get { return RubberduckUI.NonReturningFunction_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var declarations = state.AllDeclarations.ToList();

            var interfaceMembers = declarations.FindInterfaceMembers();
            var interfaceImplementationMembers = declarations.FindInterfaceImplementationMembers();

            var functions = declarations
                .Where(declaration => !declaration.IsBuiltIn && ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var issues = functions
                .Where(declaration => declaration.References.All(r => !r.IsAssignment))
                .Select(issue => new NonReturningFunctionInspectionResult(this, string.Format(Description, issue.IdentifierName), new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context), interfaceImplementationMembers.Select(m => m.Scope).Contains(issue.Scope)));

            return issues;
        }
    }
}