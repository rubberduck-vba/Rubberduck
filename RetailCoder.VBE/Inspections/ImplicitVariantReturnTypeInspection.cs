using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection(string project, string module, string procedure)
        {
            Severity = CodeInspectionSeverity.Suggestion;
            _project = project;
            _module = module;
            _procedure = procedure;
        }

        private readonly string _project;
        private readonly string _module;
        private readonly string _procedure;
        public string Name { get { return InspectionNames.ImplicitVariantReturnType; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IDictionary<QualifiedModuleName, IParseTree> nodes)
        {
            var signatures = nodes.SelectMany(node => node.Value.GetProcedures().Select(procedure => new { Key = node.Key, Procedure = procedure }))
                .Where(node => node.Procedure is VisualBasic6Parser.FunctionStmtContext
                               || node.Procedure is VisualBasic6Parser.PropertyGetStmtContext)
                .ToList();

            var functions = signatures.OfType<VisualBasic6Parser.FunctionStmtContext>()
                .Where(node => node.asTypeClause() == null)
                .Cast<ParserRuleContext>();

            var getters = signatures.OfType<VisualBasic6Parser.PropertyGetStmtContext>()
                .Where(node => node.asTypeClause() == null);

            return functions.Union(getters).Select(procedure => new ImplicitVariantReturnTypeInspectionResult(Name, procedure, Severity, _project, _module, _procedure));
        }
    }
}