using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

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
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                var declarations = (IEnumerable<ParserRuleContext>) module.ParseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener());
                foreach (var declaration in declarations.Where(declaration => declaration is VisualBasic6Parser.ConstStmtContext || declaration is VisualBasic6Parser.VariableStmtContext))
                {
                    var variables = declaration as VisualBasic6Parser.VariableStmtContext;                    
                    if (variables != null && HasMultipleDeclarations(variables))
                    {
                        yield return new MultipleDeclarationsInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, variables.variableListStmt()));
                    }

                    var consts = declaration as VisualBasic6Parser.ConstStmtContext;
                    if (consts != null && HasMultipleDeclarations(consts))
                    {
                        yield return new MultipleDeclarationsInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, consts));
                    }
                }
            }
        }

        private bool HasMultipleDeclarations(VisualBasic6Parser.VariableStmtContext context)
        {
            return context.variableListStmt().variableSubStmt().Count > 1;
        }

        private bool HasMultipleDeclarations(VisualBasic6Parser.ConstStmtContext context)
        {
            return context.constSubStmt().Count > 1;
        }
    }
}