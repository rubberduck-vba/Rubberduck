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

        private static readonly DeclarationType[] ReturningMemberTypes =
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var functions = parseResult.Declarations.Items.Where(declaration =>
                ReturningMemberTypes.Contains(declaration.DeclarationType));

            var issues = functions
                .Where(declaration => 
                    !IsInterfaceMember(parseResult.Declarations, declaration)
                    && declaration.References.All(r => !r.IsAssignment))
                .Select(issue => new NonReturningFunctionInspectionResult(string.Format(Name, issue.IdentifierName), Severity, new QualifiedContext<ParserRuleContext>(issue.QualifiedName, issue.Context)));

            return issues;
        }

        private bool IsInterfaceMember(Declarations declarations, Declaration procedure)
        {
            var parent = declarations.Items.SingleOrDefault(item =>
                        item.Project == procedure.Project &&
                        item.IdentifierName == procedure.ComponentName &&
                       (item.DeclarationType == DeclarationType.Class));

            if (parent == null)
            {
                return false;
            }

            var classes = declarations.Items.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext));

            return interfaces.Select(i => i.ComponentName).Contains(procedure.ComponentName);
        }
    }
}