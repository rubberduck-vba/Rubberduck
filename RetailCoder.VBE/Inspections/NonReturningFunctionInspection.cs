using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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

        private static readonly DeclarationType[] InterfaceMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = parseResult.Declarations.Items
                .Where(IsImplemented)
                .SelectMany(iItem => parseResult.Declarations.FindMembers(iItem)
                    .Where(member => InterfaceMemberTypes.Contains(member.DeclarationType)));

            var functions = parseResult.Declarations.Items.Where(declaration =>
                ReturningMemberTypes.Contains(declaration.DeclarationType)
                && !interfaceMembers.Contains(declaration));

            var issues = functions
                .Where(declaration => declaration.References.All(r => !r.IsAssignment))
                .Select(issue => new NonReturningFunctionInspectionResult(string.Format(Name, issue.IdentifierName), Severity, new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context)));

            return issues;
        }

        private bool IsImplemented(Declaration item)
        {
            return item.DeclarationType == DeclarationType.Class
                   && item.References.Any(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext);
        }
    }
}