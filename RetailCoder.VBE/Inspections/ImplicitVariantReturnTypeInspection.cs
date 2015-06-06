using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitVariantReturnTypeInspection : IInspection
    {
        public ImplicitVariantReturnTypeInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitVariantReturnTypeInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitVariantReturnType_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.LibraryFunction
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = from item in parseResult.Declarations.Items
                               where !item.IsBuiltIn
                                && ProcedureTypes.Contains(item.DeclarationType)
                                && !item.IsTypeSpecified()
                               let issue = new {Declaration = item, QualifiedContext = new QualifiedContext<ParserRuleContext>(item.QualifiedName, item.Context)}
                               select new ImplicitVariantReturnTypeInspectionResult(string.Format(Description, issue.Declaration.IdentifierName), Severity, issue.QualifiedContext);
            return issues;
        }

        private static readonly IEnumerable<Func<ParserRuleContext, VBAParser.AsTypeClauseContext>> Converters =
        new List<Func<ParserRuleContext, VBAParser.AsTypeClauseContext>>
            {
                GetFunctionReturnType,
                GetPropertyGetReturnType
            };

        private VBAParser.AsTypeClauseContext GetAsTypeClause(ParserRuleContext procedureContext)
        {
            return Converters.Select(converter => converter(procedureContext)).FirstOrDefault(args => args != null);
        }

        private static bool HasExpectedReturnType(QualifiedContext<ParserRuleContext> procedureContext)
        {
            var function = procedureContext.Context as VBAParser.FunctionStmtContext;
            var getter = procedureContext.Context as VBAParser.PropertyGetStmtContext;
            return function != null || getter != null;
        }

        private static VBAParser.AsTypeClauseContext GetFunctionReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBAParser.FunctionStmtContext;
            return context == null ? null : context.asTypeClause();
        }

        private static VBAParser.AsTypeClauseContext GetPropertyGetReturnType(ParserRuleContext procedureContext)
        {
            var context = procedureContext as VBAParser.PropertyGetStmtContext;
            return context == null ? null : context.asTypeClause();
        }
    }
}