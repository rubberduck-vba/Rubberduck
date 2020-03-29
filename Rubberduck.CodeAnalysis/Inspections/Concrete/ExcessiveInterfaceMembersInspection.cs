using System;
using System.Linq;
using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.SettingsProvider;

//todo:
//finish implementing settings
//implement quickfix

namespace Rubberduck.CodeAnalysis.Inspections.Concrete 
{
    /// <summary>
    /// Identifies class modules that define an interface with an excessive number of public members and reminds users about Interface Segregation Principle.
    /// </summary>
    /// <why>
    /// Interfaces should not be designed to continually grow new members; we should be keeping them small, specific, and specialized.
    /// </why>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// '@Interface
    ///
    /// Public Sub DoSomething()
    /// 
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
    /// Public Sub DoSomething1()
    /// 
    /// End Sub
    /// 
    /// Public Sub DoSomething2()
    /// 
    /// End Sub
    /// 
    /// '...
    /// 
    /// Public Sub DoSomethingNGreaterThanMaxPublicMemberCount()
    ///  
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// 
    internal sealed class ExcessiveInterfaceMembersInspection : DeclarationInspectionBase<int>
    {
        private static int PublicMemberLimit;

        //constructor for actual inspection that does not allow for changing from the default limit of 10; this should be removed when settings is fully implemented
        public ExcessiveInterfaceMembersInspection(IDeclarationFinderProvider declarationFinderProvider) 
            : base(declarationFinderProvider, DeclarationType.ClassModule)
        {
            PublicMemberLimit = 10;
        }

        //constructor only for unit test; should become only constructor once settings is fully implemented
        public ExcessiveInterfaceMembersInspection(IDeclarationFinderProvider declarationFinderProvider, IConfigurationService<int> settings) 
            : base (declarationFinderProvider, DeclarationType.ClassModule) 
        {
            PublicMemberLimit = settings.Read();
        }

        protected override (bool isResult, int properties) IsResultDeclarationWithAdditionalProperties(Declaration declaration, DeclarationFinder finder)
        {
            if (!(declaration is ClassModuleDeclaration classModule && classModule.IsInterface))
        {
                return (false, 0);
            }

            return HasExcessiveMembers(classModule);
        }

        private static (bool, int) HasExcessiveMembers(ClassModuleDeclaration declaration)
        {
            var publicMembers = declaration.Members.Where(member =>
                (member.Accessibility == Accessibility.Public) ||
                (member.Accessibility == Accessibility.Global) ||
                (member.Accessibility == Accessibility.Implicit));

            var count = publicMembers.Where(member => member.DeclarationType != DeclarationType.Event)
                            .GroupBy(member => member.IdentifierName)
                            .Select(grouping => grouping.First()).Count();

            return (count > PublicMemberLimit, count);
        }

        protected override string ResultDescription(Declaration declaration, int memberCount) 
        {
            var identifierName = declaration.IdentifierName;

            return string.Format(
                InspectionResults.ExcessiveInterfaceMembersInspection,
                identifierName,
                memberCount);
        }
    }
}