using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspection : IInspection
    {
        public ParameterNotUsedInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ParameterNotUsed_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = FindInterfaceMemberScopes(parseResult.Declarations);
            var issues = parseResult.Declarations.Items.Where(parameter =>
                parameter.DeclarationType == DeclarationType.Parameter
                && !(parameter.Context.Parent.Parent is VBAParser.EventStmtContext)
                && !(parameter.Context.Parent.Parent is VBAParser.DeclareStmtContext)
                && !interfaceMembers.Contains(parameter.ParentScope)
                && !parameter.References.Any())
            .Select(issue => new ParameterNotUsedInspectionResult(string.Format(Name, issue.IdentifierName), Severity, issue.Context, issue.QualifiedName));

            return issues;
        }

        private static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private IEnumerable<string> FindInterfaceMemberScopes(Declarations declarations)
        {
            var classes = declarations.Items.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                    reference.Context.Parent is VBAParser.ImplementsStmtContext))
                    .Select(i => i.Scope)
                    .ToList();

            return declarations.Items.Where(item => ProcedureTypes.Contains(item.DeclarationType) && interfaces.Any(i => item.ParentScope.StartsWith(i)))
                    .Select(member => member.Scope)
                    .ToList();
        }
    }
}