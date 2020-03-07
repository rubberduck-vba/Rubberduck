using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty module member blocks.
    /// </summary>
    /// <why>
    /// Methods containing no executable statements are misleading as they appear to be doing something which they actually don't.
    /// This might be the result of delaying the actual implementation for a later stage of development, and then forgetting all about that.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Sub Foo()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Sub Foo()
    ///     MsgBox "?"
    /// End Sub
    /// ]]>
    /// </example>
    internal class EmptyMethodInspection : DeclarationInspectionBase
    {
        public EmptyMethodInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Member)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration is ModuleBodyElementDeclaration member 
                   && !member.IsInterfaceMember 
                   && !member.Block.ContainsExecutableStatements();
        }

        protected override string ResultDescription(Declaration member)
        {
            var identifierName = member.IdentifierName;
            var declarationType = member.DeclarationType.ToLocalizedString();

            return string.Format(
                InspectionResults.EmptyMethodInspection,
                declarationType,
                identifierName);
        }
    }
}