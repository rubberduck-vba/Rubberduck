using Rubberduck.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Common;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies class modules that define an interface with one or more members containing a concrete implementation.
    /// </summary>
    /// <why>
    /// Interfaces provide an abstract, unified programmatic access to different objects; concrete implementations of their members should be in a separate module that 'Implements' the interface.
    /// </why>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// ' empty interface stub
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    ///     MsgBox "Hello from interface!"
    /// End Sub
    /// ]]>
    /// </example>
    internal class ImplementedInterfaceMemberInspection : DeclarationInspectionBase
    {
        public ImplementedInterfaceMemberInspection(Parsing.VBA.RubberduckParserState state)
            : base(state, DeclarationType.ClassModule)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (!IsInterfaceDeclaration(declaration))
            {
                return false;
            }

            var moduleBodyElements = finder.Members(declaration, DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>();

            return moduleBodyElements
                .Any(member => member.Block.ContainsExecutableStatements(true));
        }

        private static bool IsInterfaceDeclaration(Declaration declaration)
        {
            if (!(declaration is ClassModuleDeclaration classModule))
            {
                return false;
            }
            return classModule.IsInterface
                || declaration.Annotations.Any(an => an.Annotation is InterfaceAnnotation);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var qualifiedName = declaration.QualifiedModuleName.ToString();
            var declarationType = Resources.RubberduckUI.ResourceManager
                .GetString("DeclarationType_" + declaration.DeclarationType)
                .Capitalize();
            var identifierName = declaration.IdentifierName;

            return string.Format(
                InspectionResults.ImplementedInterfaceMemberInspection,
                qualifiedName,
                declarationType,
                identifierName);
        }
    }
}
