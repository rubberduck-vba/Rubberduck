using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Highlights implicit Public access modifiers in user code.
    /// </summary>
    /// <why>
    /// In modern VB (VB.NET), the implicit access modifier is Private, as it is in most other programming languages.
    /// Making the Public modifiers explicit can help surface potentially unexpected language defaults.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ImplicitPublicMemberInspection : DeclarationInspectionBase
    {
        public ImplicitPublicMemberInspection(RubberduckParserState state)
            : base(state, ProcedureTypes) { }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Accessibility == Accessibility.Implicit;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ImplicitPublicMemberInspection, declaration.IdentifierName);
        }
    }
}
