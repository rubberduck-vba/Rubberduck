using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public interface INewDeclarationConflictDetection : IConflictDetectionBase
    {
        bool NewDeclarationHasConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData);
        bool NewModuleDeclarationHasConflict(string name, string projectID, IConflictDetectionSessionData sessionData, out string nonConflictName);
    }

    public class NewDeclarationConflictDetection : ConflictDetectionBase, INewDeclarationConflictDetection
    {
        public NewDeclarationConflictDetection(IDeclarationFinderProvider declarationFinderProvider, /*IDeclarationProxyFactory declarationProxyFactory,*/ IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory)
        { }

        public bool NewDeclarationHasConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            if (TryResolveToConflictFreeIdentifier(proxy, sessionData))
            {
                sessionData.RegisterResolvedProxyIdentifier(proxy);
                return true;
            }
            return false;
        }

        public bool NewModuleDeclarationHasConflict(string name, string projectID, IConflictDetectionSessionData sessionData, out string nonConflictName)
        {
            nonConflictName = name;
            var hasConflict = true;
            for (var idx = 0; idx < 100 && hasConflict; idx++)
            {
                if (ModuleIdentifierMatchesProjectName(nonConflictName, projectID))
                {
                    nonConflictName = ConflictingNameModifier(nonConflictName);
                    continue;
                }

                if (ModuleIdentifierConflicts(nonConflictName, projectID).Any())
                {
                    nonConflictName = ConflictingNameModifier(nonConflictName);
                    continue;
                }
                hasConflict = false;
            }

            if (!hasConflict)
            {
                return true;
            }
            nonConflictName = string.Empty;
            return false;
        }

        private bool ModuleIdentifierMatchesProjectName(string name, string projectID)
        {
            var projectName = _declarationFinderProvider.DeclarationFinder.AllModules
                    .Where(mod => mod.ProjectId == projectID).Select(p => p.ProjectName).FirstOrDefault();

            return AreVBAEquivalent(name, projectName);
        }
    }
}
