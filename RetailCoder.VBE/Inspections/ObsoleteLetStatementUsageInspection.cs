using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ObsoleteLetStatementUsageInspection : IInspection
    {
        public ObsoleteLetStatementUsageInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ObsoleteLet; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var result in parseResult)
            {
                var module = result;
                var results = module.ParseTree.GetObsoleteStatements()
                    .OfType<VisualBasic6Parser.LetStmtContext>()
                    .Where(context => context.LET() != null && !string.IsNullOrEmpty(context.LET().GetText()))
                    .Select(context => new ObsoleteLetStatementUsageInspectionResult(Name, Severity, new QualifiedContext<VisualBasic6Parser.LetStmtContext>(module.QualifiedName, context)));
                foreach (var inspectionResult in results)
                {
                    yield return inspectionResult;
                }
            }
        }
    }
}