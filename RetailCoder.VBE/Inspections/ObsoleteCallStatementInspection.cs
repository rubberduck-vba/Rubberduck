using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteCallStatementInspection : IInspection
    {
        public ObsoleteCallStatementInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteCallStatementInspection"; } }
        public string Description { get { return RubberduckUI.ObsoleteCall; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            //note: this misses calls to procedures/functions without a Declaration object.
            // alternative is to walk the tree and listen for "CallStmt".

            var calls = (from declaration in parseResult.Declarations.Items
                from reference in declaration.References
                where (reference.Declaration.DeclarationType == DeclarationType.Function
                       || reference.Declaration.DeclarationType == DeclarationType.Procedure)
                      && reference.HasExplicitCallStatement()
                select reference).ToList();

            var issues = from reference in calls
                let context = reference.Context.Parent.Parent as VBAParser.ExplicitCallStmtContext
                where context != null
                let qualifiedContext = new QualifiedContext<VBAParser.ExplicitCallStmtContext>
                    (reference.QualifiedModuleName, (VBAParser.ExplicitCallStmtContext)reference.Context.Parent.Parent)
                select new ObsoleteCallStatementUsageInspectionResult(Description, Severity, qualifiedContext);

            return issues;
        }
    }
}