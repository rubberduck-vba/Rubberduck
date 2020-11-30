using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies public enumerations declared within worksheet modules.
    /// </summary>
    /// <why>
    /// Copying a worksheet which contains a public `Enum` declaration will duplicate the enum resulting in a state which prevents compilation.
    /// </why>
    /// <example hasResult="true">
    /// <module name="DocumentModule" type="Document Module">
    /// <![CDATA[
    /// Public Enum Foo()
    ///     ' enumeration members
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="DocumentModule" type="Document Module">
    /// <![CDATA[
    /// Private Enum Foo()
    ///     ' enumeration members
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class PublicEnumerationDeclaredWithinWorksheetInspection : DeclarationInspectionBase
    {
        public PublicEnumerationDeclaredWithinWorksheetInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider, new DeclarationType[] { DeclarationType.Enumeration })
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Public
                && declaration.QualifiedModuleName.ComponentType == ComponentType.Document;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.PublicEnumerationDeclaredWithinWorksheetInspection,
                declaration.IdentifierName,
                declaration.ParentScopeDeclaration.IdentifierName);
        }
    }
}
