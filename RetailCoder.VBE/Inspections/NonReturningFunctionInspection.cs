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

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var functions = parseResult.Declarations.Items.Where(declaration =>
                declaration.DeclarationType == DeclarationType.Function);

            var results = functions
                .Where(declaration => declaration.References.Where(r => declaration.Selection.Contains(r.Selection)).All(r => !r.IsAssignment));

            foreach (var result in results)
            {
                // todo: in Microsoft Access, this inspection should only return a result for private functions.
                //       changing an unassigned function to a "Sub" could break Access macros that reference it.
                //       doing this right may require accessing the Access object model to find usages in macros.
                yield return new NonReturningFunctionInspectionResult(string.Format(Name, result.IdentifierName), Severity, new QualifiedContext<ParserRuleContext>(result.QualifiedName, result.Context));
            }
        }
    }
}