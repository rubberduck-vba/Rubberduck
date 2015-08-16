using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class MultilineParameterInspection : IInspection 
    {
        public MultilineParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "MultilineParameterInspection"; } }
        public string Description { get { return RubberduckUI.MultilineParameter_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var multilineParameters = from p in parseResult.Declarations.Items
                .Where(item => item.DeclarationType == DeclarationType.Parameter)
                where p.Context.GetSelection().LineCount > 1
                select p;

            var issues = multilineParameters
                .Select(param => new MultilineParameterInspectionResult(string.Format(param.Context.GetSelection().LineCount > 3 ? RubberduckUI.EasterEgg_Continuator : Description, param.IdentifierName), Severity, param.Context, param.QualifiedName));

            return issues;
        }

        public class MultilineParameterInspectionResult : CodeInspectionResultBase
        {
            public MultilineParameterInspectionResult(string inspection, CodeInspectionSeverity severity, ParserRuleContext context, QualifiedMemberName qualifiedName)
                : base(inspection, severity, qualifiedName.QualifiedModuleName, context)
            {
                
            }

            public override IDictionary<string, Action> GetQuickFixes()
            {
                // todo: implement a quickfix to rewrite the signature on 1 line
                return new Dictionary<string, Action>();
            }
        }
    }
}