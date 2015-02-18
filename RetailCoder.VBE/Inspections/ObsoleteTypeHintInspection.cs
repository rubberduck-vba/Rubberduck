using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspection : IInspection
    {
        public ObsoleteTypeHintInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ObsoleteTypeHint; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            var inspectionResults = new List<CodeInspectionResultBase>();
            foreach (var module in parseResult)
            {
                var local = module;
                var declarations = module.ParseTree.GetContexts<DeclarationListener,ParserRuleContext>(new DeclarationListener());
                var results = declarations.OfType<VBParser.VariableSubStmtContext>()
                    .Where(variable => variable.TypeHint() != null
                                       && !string.IsNullOrEmpty(variable.TypeHint().GetText()))
                    .Select(variable => new { Context = variable, Hint = variable.TypeHint().GetText() })
                    .Select(result => new ObsoleteTypeHintInspectionResult(Name, Severity, new QualifiedContext<VBParser.VariableSubStmtContext>(local.QualifiedName, result.Context)));

                inspectionResults.AddRange(results);
            }

            return inspectionResults;
        }
    }
}