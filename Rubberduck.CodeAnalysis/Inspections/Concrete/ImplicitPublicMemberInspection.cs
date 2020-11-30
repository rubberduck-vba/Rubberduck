using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Highlights implicit Public access modifiers in user code.
    /// </summary>
    /// <why>
    /// In modern VB (VB.NET), the implicit access modifier is Private, as it is in most other programming languages.
    /// Making the Public modifiers explicit can help surface potentially unexpected language defaults.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplicitPublicMemberInspection : DeclarationInspectionBase
    {
        public ImplicitPublicMemberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, ProcedureTypes) { }

        private static readonly DeclarationType[] ProcedureTypes = 
        {
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
            DeclarationType.Enumeration,
            DeclarationType.UserDefinedType
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
