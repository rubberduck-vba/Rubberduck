using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies class modules that define an interface with one or more members containing a concrete implementation.
    /// </summary>
    /// <why>
    /// Interfaces provide an abstract, unified programmatic access to different objects; concrete implementations of their members should be in a separate module that 'Implements' the interface.
    /// </why>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// ' empty interface stub
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    ///     MsgBox "Hello from interface!"
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ImplementedInterfaceMemberInspection : DeclarationInspectionBase
    {
        public ImplementedInterfaceMemberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.ClassModule)
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
            var declarationType = CodeAnalysisUI.ResourceManager
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
