using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    /// <summary>
    /// An EasterEgg inspection. Because I couldn't help it.
    /// </summary>
    public class MultilineParameterInspection : IInspection 
    {
        public MultilineParameterInspection()
        {
            Severity = CodeInspectionSeverity.DoNotShow;
        }

        public string Name { get { return InspectionNames.EasterEgg; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var multilineParameters = from p in parseResult.Declarations.Items
                .Where(item => item.DeclarationType == DeclarationType.Parameter)
                where p.Context.GetSelection().LineCount > 3
                select p;

            var issues = multilineParameters
                .Select(param => new MultilineParameterInspectionResult(string.Format(Name, param.IdentifierName), Severity, param.Context, param.QualifiedName));

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
                return new Dictionary<string, Action>();
            }
        }
    }
}