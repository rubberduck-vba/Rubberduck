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
    /// <example hasResults="true">
    /// <![CDATA[
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
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
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
    /// ]]>
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
                                      declaration.Annotations.Any(pta => pta.Annotation is ObsoleteAnnotation));

            var issues = new List<IdentifierReferenceInspectionResult>();

            foreach (var declaration in declarations)
            {
                var replacementDocumentation = declaration.Annotations
                    .First(pta => pta.Annotation is ObsoleteAnnotation)
                    .AnnotationArguments
                    .FirstOrDefault() ?? string.Empty;

                issues.AddRange(declaration.References.Select(reference =>
                    new IdentifierReferenceInspectionResult(this,
                        string.Format(InspectionResults.ObsoleteMemberUsageInspection, reference.IdentifierName, replacementDocumentation),
                        State, reference)));
            }

            return issues;
        }
    }
}
