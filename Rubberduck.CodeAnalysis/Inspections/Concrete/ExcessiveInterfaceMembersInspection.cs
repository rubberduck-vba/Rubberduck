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

        public ExcessiveInterfaceMembersInspection(IDeclarationFinderProvider declarationFinderProvider) //constructor for actual inspection that does not allow for changing from the default limit of 10; this should be removed when settings is fully implemented
            : base(declarationFinderProvider, DeclarationType.ClassModule)
        {
            PublicMemberLimit = 10;
        }

        public ExcessiveInterfaceMembersInspection(IDeclarationFinderProvider declarationFinderProvider, IConfigurationService<int> settings) //constructor only for unit test; should become only constructor once settings is fully implemented
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
            var _publicmembers = declaration.Members.Where(member =>
            {
                int acc = (int)member.Accessibility;
                return acc >= (int)Accessibility.Implicit && acc <= (int)Accessibility.Global;
            });

            var count = _publicmembers.Where(member => member.DeclarationType != DeclarationType.Event)
                                  .Where(member => member.DeclarationType != DeclarationType.PropertyGet || NoMatchingSetter(member, _publicmembers))
                                  .Count();

            return (count > PublicMemberLimit, count);
        }

        private static bool NoMatchingSetter(Declaration property, IEnumerable<Declaration> members) =>
            !members.Any(member => (member.IdentifierName == property.IdentifierName) && (member != property));

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