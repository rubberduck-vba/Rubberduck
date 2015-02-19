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
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.VariableTypeNotDeclared; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var result in parseResult)
            {
                var declarations = 
                    result.ParseTree.GetContexts<
                        DeclarationListener, ParserRuleContext>(new DeclarationListener(result.QualifiedName)).ToList();
                var module = result; // to avoid access to modified closure in below lambdas

                var constants = declarations.Where(declaration => declaration.Context is VBParser.ConstSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBParser.ConstSubStmtContext>()
                                            .Where(constant => constant.AsTypeClause() == null)
                                            .Select(constant => new VariableTypeNotDeclaredInspectionResult(Name, Severity, constant, module.QualifiedName));

                var variables = declarations.Where(declaration => declaration.Context is VBParser.VariableSubStmtContext)
                                            .Select(declaration => declaration.Context)
                                            .Cast<VBParser.VariableSubStmtContext>()
                                            .Where(variable => variable.AsTypeClause() == null)
                                            .Select(variable => new VariableTypeNotDeclaredInspectionResult(Name, Severity, variable, module.QualifiedName));

                foreach (var inspectionResult in constants.Concat(variables))
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}