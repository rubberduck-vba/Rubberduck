using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Global' declaration statements.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete; use the 'Public' keyword instead.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Global Foo As Long
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Foo As Long
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObsoleteGlobalInspection : DeclarationInspectionBase
    {
        public ObsoleteGlobalInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Global 
                   && declaration.Context != null
                   && declaration.DeclarationType != DeclarationType.BracketedExpression;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                    InspectionResults.ObsoleteGlobalInspection,
                    declarationType,
                    declarationName);
        }
    }
}
