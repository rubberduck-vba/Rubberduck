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
    /// <summary>
    /// Flags usages of members marked as obsolete with an @Obsolete("justification") Rubberduck annotation.
    /// </summary>
    /// <why>
    /// Marking members as obsolete can help refactoring a legacy code base. This inspection is a tool that makes it easy to locate obsolete member calls.
    /// </why>
    /// <example>
    /// This inspection means to flag the following statement:
    /// <code>
    /// Public Sub DoSomething()
    ///     DoStuff ' member is marked as obsolete
    /// End Sub
    ///
    /// '@Obsolete("Use the newer DoThing() method instead")
    /// Private Sub DoStuff()
    ///     ' ...
    /// End Sub
    ///
    /// Private Sub DoThing()
    ///     ' ...
    /// End Sub
    /// </code>
    /// The following code should not trip this inspection:
    /// <code>
    /// Public Sub DoSomething()
    ///     DoThing
    /// End Sub
    ///
    /// '@Obsolete("Use the newer DoThing() method instead")
    /// Private Sub DoStuff()
    ///     ' ...
    /// End Sub
    ///
    /// Private Sub DoThing()
    ///     ' ...
    /// End Sub
    /// </code>
    /// </example>
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
