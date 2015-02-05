using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class MultipleDeclarationsInspection : IInspection
    {
        public MultipleDeclarationsInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.MultipleDeclarations; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VbModuleParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                var declarations = module.ParseTree.GetDeclarations();
                foreach (var declaration in declarations.Where(declaration => declaration is VisualBasic6Parser.ConstStmtContext || declaration is VisualBasic6Parser.VariableStmtContext))
                {
                    var variables = declaration as VisualBasic6Parser.VariableStmtContext;                    
                    if (variables != null && HasMultipleDeclarations(variables))
                    {
                        yield return new MultipleDeclarationsInspectionResult(Name, Severity, new QualifiedContext<VisualBasic6Parser.VariableListStmtContext>(module.QualifiedName, variables.variableListStmt()));
                    }

                    var consts = declaration as VisualBasic6Parser.ConstStmtContext;
                }
            }
        }

        private bool HasMultipleDeclarations(VisualBasic6Parser.VariableStmtContext context)
        {
            return context.ChildCount > 1;
        }

        private bool HasMultipleDeclarations(VisualBasic6Parser.ConstStmtContext context)
        {
            return context.ChildCount > 1;
        }
    }
}