using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Inspections
{
    public sealed class MultipleFolderAnnotationsInspection : InspectionBase
    {
        public MultipleFolderAnnotationsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.MultipleFolderAnnotationsInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MultipleFolderAnnotationsInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.MaintainabilityAndReadabilityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var issues = UserDeclarations.Where(declaration =>
                 (declaration.DeclarationType == DeclarationType.ClassModule
                || declaration.DeclarationType == DeclarationType.ProceduralModule)
                && declaration.Annotations.Count(annotation => annotation.AnnotationType == AnnotationType.Folder) > 1);
            return issues.Select(issue =>
                new MultipleFolderAnnotationsInspectionResult(this, issue));
        }
    }
}
