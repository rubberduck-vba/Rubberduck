using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;
using Rubberduck.VBA.ParseTreeListeners;

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
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener()).Where(HasExpectedReturnType);
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

        private static readonly IEnumerable<Func<ParserRuleContext, VBParser.AsTypeClauseContext>> Converters =
        new List<Func<ParserRuleContext, VBParser.AsTypeClauseContext>>
            {
                GetFunctionReturnType,
                GetPropertyGetReturnType
            };

        private VBParser.AsTypeClauseContext GetAsTypeClause(ParserRuleContext procedureContext)
        {
            return Converters.Select(converter => converter(procedureContext)).FirstOrDefault(args => args != null);
        }

        private static bool HasExpectedReturnType(ParserRuleContext procedureContext)
        {
            var function = procedureContext as VBParser.FunctionStmtContext;
            var getter = procedureContext as VBParser.PropertyGetStmtContext;
            return function != null || getter != null;
        }

        private static VBParser.AsTypeClauseContext GetFunctionReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.FunctionStmtContext;
            return context == null ? null : context.AsTypeClause();
        }

        private static VBParser.AsTypeClauseContext GetPropertyGetReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBParser.PropertyGetStmtContext;
            return context == null ? null : context.AsTypeClause();
        }
    }
}