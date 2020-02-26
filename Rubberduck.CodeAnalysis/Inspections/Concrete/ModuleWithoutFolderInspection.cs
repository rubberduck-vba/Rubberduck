using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a user module is missing a @Folder Rubberduck annotation.
    /// </summary>
    /// <why>
    /// Modules without a custom @Folder annotation will be grouped under the default folder in the Code Explorer toolwindow.
    /// By specifying a custom @Folder annotation, modules can be organized by functionality rather than simply listed.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Folder("Foo")
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class ModuleWithoutFolderInspection : DeclarationInspectionBase
    {
        public ModuleWithoutFolderInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Module)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return !declaration.Annotations.Any(pta => pta.Annotation is FolderAnnotation);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ModuleWithoutFolderInspection, declaration.IdentifierName);
        }
    }
}
