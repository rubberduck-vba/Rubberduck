using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags usages of members marked as obsolete with an @Obsolete("justification") Rubberduck annotation.
    /// </summary>
    /// <why>
    /// Marking members as obsolete can help refactoring a legacy code base. This inspection is a tool that makes it easy to locate obsolete member calls.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
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
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
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
    /// </module>
    /// </example>
    internal sealed class ObsoleteMemberUsageInspection : IdentifierReferenceInspectionBase
    {
        public ObsoleteMemberUsageInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var declaration = reference?.Declaration;
            return declaration != null
                   && declaration.IsUserDefined
                   && declaration.DeclarationType.HasFlag(DeclarationType.Member)
                   && declaration.Annotations.Any(pta => pta.Annotation is ObsoleteAnnotation);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var replacementDocumentation = reference.Declaration.Annotations
                                               .First(pta => pta.Annotation is ObsoleteAnnotation)
                                               .AnnotationArguments
                                               .FirstOrDefault() ?? string.Empty;
            return string.Format(
                InspectionResults.ObsoleteMemberUsageInspection, 
                reference.IdentifierName,
                replacementDocumentation);
        }
    }
}
