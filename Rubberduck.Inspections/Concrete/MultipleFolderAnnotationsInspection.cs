using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MultipleFolderAnnotationsInspection : InspectionBase
    {
        public MultipleFolderAnnotationsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
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
