using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates 'Const' declarations that are never referenced.
    /// </summary>
    /// <why>
    /// Declarations that are never used should be removed.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Const foo As Long = 42
    ///
    /// Public Sub DoSomething()
    ///     ' no reference to 'foo' anywhere...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Const foo As Long = 42
    ///
    /// Public Sub DoSomething()
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ConstantNotUsedInspection : DeclarationInspectionBase
    {
        public ConstantNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Constant)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration?.Context != null 
                   && !declaration.References.Any()
                   && !IsPublicInExposedClass(declaration);
        }

        private static bool IsPublicInExposedClass(Declaration procedure)
        {
            if (!(procedure.Accessibility == Accessibility.Public
                    || procedure.Accessibility == Accessibility.Global))
            {
                return false;
            }

            if (!(Declaration.GetModuleParent(procedure) is ClassModuleDeclaration classParent))
            {
                return false;
            }

            return classParent.IsExposed;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.IdentifierNotUsedInspection,
                declarationType,
                declarationName);
        }
    }
}
