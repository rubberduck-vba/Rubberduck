using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.Inspections
{
    //public class VariableNotAssignedInspection : IInspection
    //{
    //    public VariableNotAssignedInspection()
    //    {
    //        Severity = CodeInspectionSeverity.Error;
    //    }

    //    public string Name { get { return InspectionNames.VariableNotAssigned; } }
    //    public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
    //    public CodeInspectionSeverity Severity { get; set; }

    //    public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
    //    {
    //        var parseResults = parseResult.ToList();
            
    //        /* 
    //         * 
    //         * 
    //         */

    //        var assignments = parseResults.Select(result =>
    //            new
    //            {
    //                Module = result.Component,
    //                ParseTree = result.ParseTree.GetContexts<VariableAssignmentListener, VisualBasic6Parser.VariableCallStmtContext>(new VariableAssignmentListener())
    //            }).ToList();

    //        foreach (var result in parseResults)
    //        {
    //            var module = result; // to avoid access to modified closure in below lambdas
    //            var declarations = result.ParseTree.GetContexts<DeclarationListener, ParserRuleContext>(new DeclarationListener()).ToList();
    //            var localAssigns = assignments.Where(assign => assign.Module.QualifiedName().Equals(module.QualifiedName));

    //            var variables = declarations.Where(declaration => declaration is VisualBasic6Parser.VariableSubStmtContext)
    //                .Cast<VisualBasic6Parser.VariableSubStmtContext>()
    //                .Where(variable => assignments.Where(a => a.Module.QualifiedName() != ).All(assigned => assigned.ambiguousIdentifier().GetText() != variable.ambiguousIdentifier().GetText()))
    //                .Select(variable => new VariableNotAssignedInspetionResult(Name, Severity, variable, module.QualifiedName));

    //            foreach (var variable in variables)
    //            {
    //                yield return variable;
    //            }
    //        }
    //    }
    //}
}