using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies auto-assigned object declarations.
    /// </summary>
    /// <why>
    /// Auto-assigned objects are automatically re-created as soon as they are referenced. It is therefore impossible to set one such reference 
    /// to 'Nothing' and then verifying whether the object 'Is Nothing': it will never be. This behavior is potentially confusing and bug-prone.
    /// </why>
    /// <example hasResults="true">
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
    /// </example>
    /// <example hasResults="false">
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
    /// </example>
    public sealed class SelfAssignedDeclarationInspection : InspectionBase
    {
        public SelfAssignedDeclarationInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration => declaration.IsSelfAssigned 
                    && declaration.IsTypeSpecified
                    && !SymbolList.ValueTypes.Contains(declaration.AsTypeName)
                    && (declaration.AsTypeDeclaration == null
                        || declaration.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)
                    && declaration.ParentScopeDeclaration != null
                    && declaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionResults.SelfAssignedDeclarationInspection, issue.IdentifierName),
                                                      issue));
        }
    }
}
