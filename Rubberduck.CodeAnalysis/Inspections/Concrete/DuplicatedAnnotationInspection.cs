using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class DuplicatedAnnotationInspection : InspectionBase
    {
        public DuplicatedAnnotationInspection(RubberduckParserState state) : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var issues = new List<DeclarationInspectionResult>();

            foreach (var declaration in State.AllUserDeclarations)
            {
                var duplicateAnnotations = declaration.Annotations
                    .GroupBy(annotation => annotation.AnnotationType)
                    .Where(group => !group.First().AllowMultiple && group.Count() > 1);

                issues.AddRange(duplicateAnnotations.Select(duplicate =>
                {
                    var result = new DeclarationInspectionResult(
                        this,
                        string.Format(InspectionResults.DuplicatedAnnotationInspection, duplicate.Key.ToString()),
                        declaration);

                    result.Properties.AnnotationType = duplicate.Key;

                    return result;
                }));
            }

            return issues;
        }
    }
}
