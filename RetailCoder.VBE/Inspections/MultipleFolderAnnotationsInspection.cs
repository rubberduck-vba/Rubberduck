using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class MultipleFolderAnnotationsInspection : InspectionBase
    {
        public MultipleFolderAnnotationsInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return InspectionsUI.MultipleFolderAnnotationsInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations.Where(declaration =>
                 (declaration.DeclarationType == DeclarationType.Class
                || declaration.DeclarationType == DeclarationType.Module)
                && declaration.Annotations.Split('\n').Count(annotation =>
                    annotation.StartsWith(Parsing.Grammar.Annotations.AnnotationMarker +
                                          Parsing.Grammar.Annotations.Folder)) > 1);
            return issues.Select(issue =>
                new MultipleFolderAnnotationsInspectionResult(this, string.Format(Description, issue.ComponentName), issue));
        }
    }

    public class MultipleFolderAnnotationsInspectionResult : CodeInspectionResultBase
    {
        public MultipleFolderAnnotationsInspectionResult(IInspection inspection, string result, Declaration target) 
            : base(inspection, result, target)
        {
        }
    }
}