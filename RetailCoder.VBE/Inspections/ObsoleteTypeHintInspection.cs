using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspection : IInspection
    {
        public ObsoleteTypeHintInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return InspectionNames.ObsoleteTypeHint_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = from item in parseResult.Declarations.Items
                where item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(string.Format(Name, item.IdentifierName), Severity, new QualifiedContext(item.QualifiedName, item.Context));

            var references = from item in parseResult.Declarations.Items.SelectMany(d => d.References)
                where item.HasTypeHint()
                select new ObsoleteTypeHintInspectionResult(string.Format(Name, item.IdentifierName), Severity, new QualifiedContext(item.QualifiedModuleName, item.Context));

            return declarations.Union(references);
        }
    }
}