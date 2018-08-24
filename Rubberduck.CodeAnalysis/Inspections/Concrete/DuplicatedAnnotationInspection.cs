using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
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
                    .Where(group => !group.First().AllowMultiple && group.Count() > 1 &&
                                    (group.Key.HasFlag(AnnotationType.ModuleAnnotation) &&
                                     declaration.DeclarationType.HasFlag(DeclarationType.Module) ||
                                     group.Key.HasFlag(AnnotationType.MemberAnnotation) &&
                                     declaration.DeclarationType.HasFlag(DeclarationType.Member)));

                issues.AddRange(duplicateAnnotations.Select(duplicate => new DeclarationInspectionResult(
                    this,
                    string.Format(InspectionResults.DuplicatedAnnotationInspection, duplicate.Key.ToString()),
                    declaration)));
            }

            return issues;
        }
    }
}
