using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class MultipleFolderAnnotationsInspection : InspectionBase
    {
        public MultipleFolderAnnotationsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.MultipleFolderAnnotationsInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MultipleFolderAnnotationsInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations.Where(declaration =>
                 (declaration.DeclarationType == DeclarationType.Class
                || declaration.DeclarationType == DeclarationType.Module)
                && declaration.Annotations.Split('\n').Count(annotation =>
                    annotation.StartsWith(Parsing.Grammar.Annotations.AnnotationMarker +
                                          Parsing.Grammar.Annotations.Folder)) > 1);
            return issues.Select(issue =>
                new MultipleFolderAnnotationsInspectionResult(this, issue));
        }
    }

    public class MultipleFolderAnnotationsInspectionResult : InspectionResultBase
    {
        public MultipleFolderAnnotationsInspectionResult(IInspection inspection, Declaration target) 
            : base(inspection, target)
        {
        }

        public override string Description
        {
            get { return string.Format(Inspection.Description, Target.IdentifierName); }
        }
    }
}