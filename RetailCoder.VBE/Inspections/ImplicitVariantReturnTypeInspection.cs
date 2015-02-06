using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ImplicitVariantReturnType; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IEnumerable<VBComponentParseResult> parseResult)
        {
            foreach (var module in parseResult)
            {
                // bug: module.ParseTree.GetProcedures() returns 0 items, listener doesn't seem to work??
                var procedures = module.ParseTree.GetProcedures().Where(HasExpectedReturnType);
                foreach (var procedure in procedures)
                {
                    var asTypeClause = GetAsTypeClause(procedure);
                    if (asTypeClause == null)
                    {
                        yield return new ImplicitVariantReturnTypeInspectionResult(Name, Severity, 
                            new QualifiedContext<ParserRuleContext>(module.QualifiedName, procedure));
                    }
                }
            }
        }

        private static readonly IEnumerable<Func<ParserRuleContext, VisualBasic6Parser.AsTypeClauseContext>> Converters =
        new List<Func<ParserRuleContext, VisualBasic6Parser.AsTypeClauseContext>>
            {
                GetFunctionReturnType,
                GetPropertyGetReturnType
            };

        private VisualBasic6Parser.AsTypeClauseContext GetAsTypeClause(ParserRuleContext procedureContext)
        {
            return Converters.Select(converter => converter(procedureContext)).FirstOrDefault(args => args != null);
        }

        private static bool HasExpectedReturnType(ParserRuleContext procedureContext)
        {
            var function = procedureContext as VisualBasic6Parser.FunctionStmtContext;
            var getter = procedureContext as VisualBasic6Parser.PropertyGetStmtContext;
            return function != null || getter != null;
        }

        private static VisualBasic6Parser.AsTypeClauseContext GetFunctionReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.FunctionStmtContext;
            return context == null ? null : context.asTypeClause();
        }

        private static VisualBasic6Parser.AsTypeClauseContext GetPropertyGetReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VisualBasic6Parser.PropertyGetStmtContext;
            return context == null ? null : context.asTypeClause();
        }
    }
}