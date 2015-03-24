using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Listeners;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;

namespace Rubberduck.Inspections
{
    public class MultipleDeclarationsInspection : IInspection
    {
        public MultipleDeclarationsInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.MultipleDeclarations; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var module in parseResult.ComponentParseResults)
            {
                var declarations = module.ParseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener(module.QualifiedName));
                foreach (var declaration in declarations.Where(declaration => declaration.Context is VBAParser.ConstStmtContext || declaration.Context is VBAParser.VariableStmtContext))
                {
                    var variables = declaration.Context as VBAParser.VariableStmtContext;                    
                    if (variables != null && HasMultipleDeclarations(variables))
                    {
                        yield return new MultipleDeclarationsInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, variables.variableListStmt()));
                    }

                    var consts = declaration.Context as VBAParser.ConstStmtContext;
                    if (consts != null && HasMultipleDeclarations(consts))
                    {
                        yield return new MultipleDeclarationsInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, consts));
                    }
                }
            }
        }

        private bool HasMultipleDeclarations(VBAParser.VariableStmtContext context)
        {
            return context.variableListStmt().variableSubStmt().Count > 1;
        }

        private bool HasMultipleDeclarations(VBAParser.ConstStmtContext context)
        {
            return context.constSubStmt().Count > 1;
        }
    }
}