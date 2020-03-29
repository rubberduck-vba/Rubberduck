using System.Linq;
using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Common;
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
    internal sealed class ImplementedInterfaceMemberInspection : DeclarationInspectionBase<IEnumerable<ModuleBodyElementDeclaration>>
    {
        public ImplementedInterfaceMemberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.ClassModule) 
        {}

        protected override (bool, IEnumerable<ModuleBodyElementDeclaration>) IsResultDeclarationWithAdditionalProperties(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ClassModuleDeclaration classModule && classModule.IsInterface)) 
            {
                return (false, null);
            }

            var implementedMembers = FindImplementedMembers(declaration, finder);
            return (implementedMembers.Count() > 0, implementedMembers);
        }

        private IEnumerable<ModuleBodyElementDeclaration> FindImplementedMembers(Declaration declaration, DeclarationFinder finder) 
        {
            var moduleBodyElements = finder.Members(declaration, DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>();

            return moduleBodyElements
                .Where(member => member.Block.ContainsExecutableStatements(true));
        }

        protected override string ResultDescription(Declaration declaration, IEnumerable<ModuleBodyElementDeclaration> results)
        {
            var identifierName = declaration.IdentifierName;
            var memberResultsString = FormatConcreteImplementationsList(results);

            return string.Format(
                InspectionResults.ImplementedInterfaceMemberInspection,
                identifierName,
                memberResultsString);
        }

        private static string FormatConcreteImplementationsList(IEnumerable<ModuleBodyElementDeclaration> results)
        {
            var items = results.Select(result => $"{result.IdentifierName} ({Resources.RubberduckUI.ResourceManager.GetString("DeclarationType_" + result.DeclarationType).Capitalize()})");
            return string.Join(", ", items);
        }
    }
}