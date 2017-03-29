using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteTypeHintInspection : InspectionBase
    {
        public ObsoleteTypeHintInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var results = UserDeclarations.ToList();

            var declarations = from item in results
                where item.HasTypeHint
                select
                    new ObsoleteTypeHintInspectionResult(this,
                        string.Format(InspectionsUI.ObsoleteTypeHintInspectionResultFormat,
                            InspectionsUI.Inspections_Declaration, item.DeclarationType.ToString().ToLower(),
                            item.IdentifierName), new QualifiedContext(item.QualifiedName, item.Context), item);

            var references = from item in results.SelectMany(d => d.References)
                where item.HasTypeHint()
                select
                    new ObsoleteTypeHintInspectionResult(this,
                        string.Format(InspectionsUI.ObsoleteTypeHintInspectionResultFormat,
                            InspectionsUI.Inspections_Usage, item.Declaration.DeclarationType.ToString().ToLower(),
                            item.IdentifierName), new QualifiedContext(item.QualifiedModuleName, item.Context),
                        item.Declaration);

            return declarations.Union(references);
        }
    }
}
