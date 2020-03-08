using System.Collections.Generic;
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
    /// Warns about duplicated annotations.
    /// </summary>
    /// <why>
    /// Rubberduck annotations should not be specified more than once for a given module, member, variable, or expression.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Bar")
    /// '@Folder("Foo")
    ///
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Foo.Bar")
    ///
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class DuplicatedAnnotationInspection : DeclarationInspectionMultiResultBase<IAnnotation>
    {
        public DuplicatedAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider) 
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IAnnotation> ResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Annotations
                .GroupBy(pta => pta.Annotation)
                .Where(group => !group.First().Annotation.AllowMultiple && group.Count() > 1)
                .Select(group => group.Key);
        }

        protected override string ResultDescription(Declaration declaration, IAnnotation annotation)
        {
            return string.Format(InspectionResults.DuplicatedAnnotationInspection, annotation);
        }
    }
}
