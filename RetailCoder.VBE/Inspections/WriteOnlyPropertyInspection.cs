using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class WriteOnlyPropertyInspection : IInspection
    {
        public WriteOnlyPropertyInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "WriteOnlyPropertyInspection"; } }
        public string Description { get { return RubberduckUI.WriteOnlyProperty_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(IRubberduckParserState parseResult)
        {
            var declarations = parseResult.AllDeclarations.ToList();
            var setters = declarations
                .Where(item => !item.IsBuiltIn 
                    && (item.Accessibility == Accessibility.Implicit || 
                        item.Accessibility == Accessibility.Public || 
                        item.Accessibility == Accessibility.Global)
                    && (item.DeclarationType == DeclarationType.PropertyLet ||
                        item.DeclarationType == DeclarationType.PropertySet)
                    && !declarations.Where(declaration => declaration.IdentifierName == item.IdentifierName)
                        .Any(accessor => !accessor.IsBuiltIn && accessor.DeclarationType == DeclarationType.PropertyGet));

            //note: if property has both Set and Let accessors, this generates 2 results.
            return setters.Select(setter =>
                new WriteOnlyPropertyInspectionResult(this, string.Format(Description, setter.IdentifierName), setter));
        }
    }

    public class WriteOnlyPropertyInspectionResult : CodeInspectionResultBase
    {
        public WriteOnlyPropertyInspectionResult(IInspection inspection, string result, Declaration target) 
            : base(inspection, result, target)
        {
        }

        // todo: override quickfixes
        //public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get; private set; }
    }
}