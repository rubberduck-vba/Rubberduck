using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates module-level fields that can be moved to a smaller scope.
    /// </summary>
    /// <why>
    /// Module-level variables that are only used in a single procedure can often be declared in that procedure's scope. 
    /// Declaring variables closer to where they are used generally makes the code easier to follow.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// Private foo As Long
    ///
    /// Public Sub DoSomething()
    ///     foo = 42
    ///     Debug.Print foo ' module variable is only used in this scope
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    ///     Dim foo As Long ' local variable only used in this scope
    ///     foo = 42
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class MoveFieldCloserToUsageInspection : InspectionBase
    {
        public MoveFieldCloserToUsageInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(declaration =>
                {
                    if (declaration.IsWithEvents
                        || !new[] {DeclarationType.ClassModule, DeclarationType.Document, DeclarationType.ProceduralModule}.Contains(declaration.ParentDeclaration.DeclarationType))
                    {
                        return false;
                    }

                    var asType = declaration.AsTypeDeclaration;
                    if (asType != null && asType.ProjectName.Equals("Rubberduck") &&
                        (asType.IdentifierName.Equals("PermissiveAssertClass") || asType.IdentifierName.Equals("AssertClass")))
                    {
                        return false;
                    }

                    var firstReference = declaration.References.FirstOrDefault();

                    if (firstReference == null ||
                        declaration.References.Any(r => r.ParentScoping != firstReference.ParentScoping))
                    {
                        return false;
                    }

                    var parentDeclaration = ParentDeclaration(firstReference);

                    return parentDeclaration != null &&
                           !new[]
                           {
                               DeclarationType.PropertyGet,
                               DeclarationType.PropertyLet,
                               DeclarationType.PropertySet
                           }.Contains(parentDeclaration.DeclarationType);
                })
                .Select(issue =>
                        new DeclarationInspectionResult(this, string.Format(InspectionResults.MoveFieldCloserToUsageInspection, issue.IdentifierName), issue));
        }

        private Declaration ParentDeclaration(IdentifierReference reference)
        {
            var declarationTypes = new[] {DeclarationType.Function, DeclarationType.Procedure, DeclarationType.Property};

            return UserDeclarations.SingleOrDefault(d =>
                        reference.ParentScoping.Equals(d) && declarationTypes.Contains(d.DeclarationType) &&
                        d.QualifiedName.QualifiedModuleName.Equals(reference.QualifiedModuleName));
        }
    }
}
