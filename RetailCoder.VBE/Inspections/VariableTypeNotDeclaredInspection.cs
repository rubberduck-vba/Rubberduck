using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Listeners;

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

                var constants = declarations.Where(declaration => declaration.Context is VBAParser.ConstSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBAParser.ConstSubStmtContext>()
                                            .Where(constant => constant.asTypeClause() == null && constant.typeHint() == null)
                                            .Select(constant => new VariableTypeNotDeclaredInspectionResult(string.Format(Name, constant.ambiguousIdentifier().GetText()), Severity, constant, module.QualifiedName));

                var variables = declarations.Where(declaration => declaration.Context is VBAParser.VariableSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBAParser.VariableSubStmtContext>()
                                            .Where(variable => variable.asTypeClause() == null && variable.typeHint() == null)
                                            .Select(variable => new VariableTypeNotDeclaredInspectionResult(string.Format(Name, variable.ambiguousIdentifier().GetText()), Severity, variable, module.QualifiedName));

                foreach (var inspectionResult in constants.Concat(variables))
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}