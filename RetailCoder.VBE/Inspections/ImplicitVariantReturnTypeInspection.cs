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
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return InspectionNames.ImplicitVariantReturnType_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            foreach (var module in parseResult.ComponentParseResults)
            {
                var procedures = module.ParseTree.GetContexts<ProcedureListener, ParserRuleContext>(new ProcedureListener(module.QualifiedName))
                    .Where(HasExpectedReturnType);
                foreach (var procedure in procedures)
                {
                    var asTypeClause = GetAsTypeClause(procedure.Context);
                    if (asTypeClause == null)
                    {
                        yield return new ImplicitVariantReturnTypeInspectionResult(string.Format(Name, ((dynamic)procedure.Context).AmbiguousIdentifier().GetText()), Severity, procedure);
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

        private static bool HasExpectedReturnType(QualifiedContext<ParserRuleContext> procedureContext)
        {
            var function = procedureContext.Context as VBParser.FunctionStmtContext;
            var getter = procedureContext.Context as VBParser.PropertyGetStmtContext;
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