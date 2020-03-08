using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates module-level fields that can be moved to a smaller scope.
    /// </summary>
    /// <why>
    /// Module-level variables that are only used in a single procedure can often be declared in that procedure's scope. 
    /// Declaring variables closer to where they are used generally makes the code easier to follow.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private foo As Long
    ///
    /// Public Sub DoSomething()
    ///     foo = 42
    ///     Debug.Print foo ' module variable is only used in this scope
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    ///
    /// Public Sub DoSomething()
    ///     Dim foo As Long ' local variable only used in this scope
    ///     foo = 42
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MoveFieldCloserToUsageInspection : DeclarationInspectionBase
    {
        public MoveFieldCloserToUsageInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Variable)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (declaration.IsWithEvents
                || !IsField(declaration))
            {
                return false;
            }

            if (IsRubberduckAssertField(declaration))
            {
                return false;
            }

            var firstReference = declaration.References.FirstOrDefault();
            var usageMember = firstReference?.ParentScoping;

            if (usageMember == null 
                || declaration.References.Any(reference => !reference.ParentScoping.Equals(usageMember)))
            {
                return false;
            }

            return usageMember.DeclarationType == DeclarationType.Procedure
                   || usageMember.DeclarationType == DeclarationType.Function;
        }

        private static bool IsField(Declaration variableDeclaration)
        {
            var parentDeclarationType = variableDeclaration.ParentDeclaration.DeclarationType;
            return parentDeclarationType.HasFlag(DeclarationType.Module);
        }

        private static bool IsRubberduckAssertField(Declaration fieldDeclaration)
        {
            var asType = fieldDeclaration.AsTypeDeclaration;
            return asType != null
                   && asType.ProjectName.Equals("Rubberduck")
                   && (asType.IdentifierName.Equals("PermissiveAssertClass")
                       || asType.IdentifierName.Equals("AssertClass"));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.MoveFieldCloserToUsageInspection, declaration.IdentifierName);
        }
    }
}
