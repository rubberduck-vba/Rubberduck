using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about duplicated annotations.
    /// </summary>
    /// <why>
    /// Rubberduck annotations should not be specified more than once for a given module, member, variable, or expression.
    /// </why>
    /// <example>
    /// This inspection means to flag the following examples:
    /// <code>
    /// '@Folder("Bar")
    /// '@Folder("Foo")
    ///
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// </code>
    /// The following code should not trip this inspection:
    /// <code>
    /// '@Folder("Foo.Bar")
    ///
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// </code>
    /// </example>
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
                        this, string.Format(InspectionResults.DuplicatedAnnotationInspection, duplicate.Key.ToString()), declaration);

                    result.Properties.AnnotationType = duplicate.Key;
                    return result;
                }));
            }

            return issues;
        }
    }
}
