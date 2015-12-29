using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspection : IInspection
    {
        public ObsoleteTypeHintInspection()
        {
            Severity = CodeInspectionSeverity.Suggestion;
        }

        public string Name { get { return "ObsoleteTypeHintInspection"; } }
        public string Meta { get { return InspectionsUI.ResourceManager.GetString(Name + "Meta"); } }
        public string Description { get { return RubberduckUI._ObsoleteTypeHint_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(RubberduckParserState state)
        {
            var results = state.AllUserDeclarations.ToList();

            var declarations = from item in results
                where item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(this, string.Format(Description, RubberduckUI.Inspections_DeclarationOf + item.DeclarationType.ToString().ToLower(), item.IdentifierName), new QualifiedContext(item.QualifiedName, item.Context), item);

            var references = from item in results.SelectMany(d => d.References)
                where item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(this, string.Format(Description, RubberduckUI.Inspections_UsageOf + item.Declaration.DeclarationType.ToString().ToLower(), item.IdentifierName), new QualifiedContext(item.QualifiedModuleName, item.Context), item.Declaration);

            return declarations.Union(references);
        }
    }
}