using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspection : IInspection
    {
        public ImplicitByRefParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return InspectionNames.ImplicitByRef_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var interfaceMembers = parseResult.Declarations.FindInterfaceImplementationMembers();
            var issues = (from item in parseResult.Declarations.Items
                where item.DeclarationType == DeclarationType.Parameter
                    && !item.IsBuiltIn
                    && !interfaceMembers.Select(m => m.Scope).Contains(item.ParentScope)
                let arg = item.Context.Parent as VBAParser.ArgContext
                where arg != null && arg.BYREF() == null && arg.BYVAL() == null
                select new QualifiedContext<VBAParser.ArgContext>(item.QualifiedName, arg))
                .Select(issue => new ImplicitByRefParameterInspectionResult(string.Format(Name, issue.Context.ambiguousIdentifier().GetText()), Severity, issue));

            return issues;
        }
    }
}