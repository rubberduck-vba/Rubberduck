using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    public class NonReturningFunctionInspection : IInspection
    {
        public NonReturningFunctionInspection()
        {
            Severity = CodeInspectionSeverity.Error;
        }

        public string Name { get { return InspectionNames.NonReturningFunction; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener());
                var functions = procedures.OfType<VisualBasic6Parser.FunctionStmtContext>()
                    .Where(function => function.GetContexts<VariableAssignmentListener, VisualBasic6Parser.AmbiguousIdentifierContext>(new VariableAssignmentListener())
                        .All(assignment => assignment.GetText() != function.ambiguousIdentifier().GetText()));
                foreach (var unassignedFunction in functions)
                {
                    yield return new NonReturningFunctionInspectionResult(Name, Severity, new QualifiedContext<ParserRuleContext>(module.QualifiedName, unassignedFunction));
                }
            }
        }
    }
}