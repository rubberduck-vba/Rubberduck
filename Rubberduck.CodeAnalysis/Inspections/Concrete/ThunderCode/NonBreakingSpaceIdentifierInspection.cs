using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates non-breaking spaces hidden in identifier names.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// This inspection may accidentally reveal non-breaking spaces in code copied and pasted from a website.
    /// </why>
    internal sealed class NonBreakingSpaceIdentifierInspection : DeclarationInspectionBase
    {
        private const string Nbsp = "\u00A0";

        public NonBreakingSpaceIdentifierInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IdentifierName.Contains(Nbsp);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return InspectionResults.NonBreakingSpaceIdentifierInspection.ThunderCodeFormat(declaration.IdentifierName);
        }
    }
}
