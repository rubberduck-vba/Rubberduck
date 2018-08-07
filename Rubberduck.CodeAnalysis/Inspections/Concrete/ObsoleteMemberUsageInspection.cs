using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    public sealed class ObsoleteMemberUsageInspection : InspectionBase
    {
        public ObsoleteMemberUsageInspection(RubberduckParserState state) : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarations = State.AllUserDeclarations
                .Where(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Member) &&
                                      declaration.Annotations.Any(annotation =>annotation.AnnotationType == AnnotationType.Obsolete));

            var issues = new List<IdentifierReferenceInspectionResult>();

            foreach (var declaration in declarations)
            {
                var replacementDocumentation =
                ((ObsoleteAnnotation) declaration.Annotations.First(annotation =>
                    annotation.AnnotationType == AnnotationType.Obsolete)).ReplacementDocumentation;

                issues.AddRange(declaration.References.Select(reference =>
                    new IdentifierReferenceInspectionResult(this,
                        string.Format(InspectionResults.ObsoleteMemberUsageInspection, reference.IdentifierName, replacementDocumentation),
                        State, reference)));
            }

            return issues;
        }
    }
}
