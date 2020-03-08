using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies auto-assigned object declarations.
    /// </summary>
    /// <why>
    /// Auto-assigned objects are automatically re-created as soon as they are referenced. It is therefore impossible to set one such reference 
    /// to 'Nothing' and then verifying whether the object 'Is Nothing': it will never be. This behavior is potentially confusing and bug-prone.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim c As New Collection
    ///     Set c = Nothing
    ///     c.Add 42 ' no error 91 raised
    ///     Debug.Print c.Count ' 1
    ///     Set c = Nothing
    ///     Debug.Print c Is Nothing ' False
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim c As Collection
    ///     Set c = New Collection
    ///     Set c = Nothing
    ///     c.Add 42 ' error 91
    ///     Debug.Print c.Count ' error 91
    ///     Set c = Nothing
    ///     Debug.Print c Is Nothing ' True
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class SelfAssignedDeclarationInspection : DeclarationInspectionBase
    {
        public SelfAssignedDeclarationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Variable)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IsSelfAssigned
                   && declaration.IsTypeSpecified
                   && !SymbolList.ValueTypes.Contains(declaration.AsTypeName)
                   && (declaration.AsTypeDeclaration == null
                       || declaration.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)
                   && declaration.ParentScopeDeclaration != null
                   && declaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.SelfAssignedDeclarationInspection, declaration.IdentifierName);
        }
    }
}
