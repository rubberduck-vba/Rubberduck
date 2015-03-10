using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
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

        public string Name { get { return InspectionNames.ObsoleteTypeHint_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var inspectionResults = new List<CodeInspectionResultBase>();
            foreach (var module in parseResult.ComponentParseResults)
            {
                var local = module;
                var declarations = module.ParseTree.GetContexts<DeclarationListener,ParserRuleContext>(new DeclarationListener(module.QualifiedName));
                var results = declarations.Select(declaration => declaration.Context).OfType<VBAParser.VariableSubStmtContext>()
                    .Where(variable => variable.typeHint() != null
                                       && !string.IsNullOrEmpty(variable.typeHint().GetText()))
                    .Select(variable => new { Context = variable, Hint = variable.typeHint().GetText() })
                    .Select(result => new ObsoleteTypeHintInspectionResult(string.Format(Name, result.Context.ambiguousIdentifier().GetText()), Severity, new QualifiedContext<VBAParser.VariableSubStmtContext>(local.QualifiedName, result.Context)));

                inspectionResults.AddRange(results);
            }

            return inspectionResults;
        }
    }
}