using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Global' declaration statements.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete; use the 'Public' keyword instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Global Foo As Long
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Foo As Long
    /// ]]>
    /// </example>
    public sealed class ObsoleteGlobalInspection : DeclarationInspectionBase
    {
        public ObsoleteGlobalInspection(RubberduckParserState state)
            : base(state) { }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Global 
                   && declaration.Context != null;
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
