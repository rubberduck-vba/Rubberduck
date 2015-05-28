using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspection : IInspection
    {
        public ObsoleteTypeHintInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return RubberduckUI._ObsoleteTypeHint_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = from item in parseResult.Declarations.Items
                where !item.IsBuiltIn && item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(string.Format(Name, RubberduckUI.Inspections_DeclarationOf + item.DeclarationType.ToString().ToLower(), item.IdentifierName), Severity, new QualifiedContext(item.QualifiedName, item.Context), item);

            var references = from item in parseResult.Declarations.Items.Where(item => !item.IsBuiltIn).SelectMany(d => d.References)
                where item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(string.Format(Name, RubberduckUI.Inspections_UsageOf + item.Declaration.DeclarationType.ToString().ToLower(), item.IdentifierName), Severity, new QualifiedContext(item.QualifiedModuleName, item.Context), item.Declaration);

            return declarations.Union(references);
        }
    }
}