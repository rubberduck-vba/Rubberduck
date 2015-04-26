using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspection : IInspection
    {
        public NonReturningFunctionInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.NonReturningFunction_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = parseResult.Declarations.FindInterfaceMembers();
            var interfaceImplementationMembers = parseResult.Declarations.FindInterfaceImplementationMembers();

            var functions = parseResult.Declarations.Items
                .Where(declaration => ReturningMemberTypes.Contains(declaration.DeclarationType)
                    && !interfaceMembers.Contains(declaration)).ToList();

            var issues = functions
                .Where(declaration => declaration.References.All(r => !r.IsAssignment))
                .Select(issue => new NonReturningFunctionInspectionResult(string.Format(Name, issue.IdentifierName), Severity, new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context), interfaceImplementationMembers.Select(m => m.Scope).Contains(issue.Scope)));

            return issues;
        }
    }
}