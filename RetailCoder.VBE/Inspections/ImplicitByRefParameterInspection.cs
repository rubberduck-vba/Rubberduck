using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ImplicitByRefParameterInspection : IInspection
    {
        public ImplicitByRefParameterInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "ImplicitByRefParameterInspection"; } }
        public string Description { get { return RubberduckUI.ImplicitByRef_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        private string AnnotationName { get { return Name.Replace("Inspection", string.Empty); } }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState parseResult)
        {
            var declarations = parseResult.AllDeclarations.ToList();

            var interfaceMembers = declarations.FindInterfaceImplementationMembers();

            var issues = (from item in declarations
                where !item.IsInspectionDisabled(AnnotationName)
                    && item.DeclarationType == DeclarationType.Parameter
                    && !item.IsBuiltIn
                    && !interfaceMembers.Select(m => m.Scope).Contains(item.ParentScope)
                let arg = item.Context as VBAParser.ArgContext
                where arg != null && arg.BYREF() == null && arg.BYVAL() == null
                select new QualifiedContext<VBAParser.ArgContext>(item.QualifiedName, arg))
                .Select(issue => new ImplicitByRefParameterInspectionResult(this, string.Format(Description, issue.Context.ambiguousIdentifier().GetText()), issue));

            return issues;
        }
    }
}