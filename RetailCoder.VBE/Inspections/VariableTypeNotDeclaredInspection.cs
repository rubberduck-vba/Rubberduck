using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class VariableTypeNotDeclaredInspection : IInspection
    {
        public VariableTypeNotDeclaredInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.VariableTypeNotDeclared_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var result in parseResult.ComponentParseResults)
            {
                var declarations = 
                    result.ParseTree.GetContexts<
                        DeclarationListener, ParserRuleContext>(new DeclarationListener(result.QualifiedName)).ToList();
                var module = result; // to avoid access to modified closure in below lambdas

                // we want declarations with a null AsTypeClause() and a null TypeHint().

                var constants = declarations.Where(declaration => declaration.Context is VBParser.ConstSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBParser.ConstSubStmtContext>()
                                            .Where(constant => constant.AsTypeClause() == null && constant.TypeHint() == null)
                                            .Select(constant => new VariableTypeNotDeclaredInspectionResult(string.Format(Name, constant.AmbiguousIdentifier().GetText()), Severity, constant, module.QualifiedName));

                var variables = declarations.Where(declaration => declaration.Context is VBParser.VariableSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBParser.VariableSubStmtContext>()
                                            .Where(variable => variable.AsTypeClause() == null && variable.TypeHint() == null)
                                            .Select(variable => new VariableTypeNotDeclaredInspectionResult(string.Format(Name, variable.AmbiguousIdentifier().GetText()), Severity, variable, module.QualifiedName));

                foreach (var inspectionResult in constants.Concat(variables))
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}