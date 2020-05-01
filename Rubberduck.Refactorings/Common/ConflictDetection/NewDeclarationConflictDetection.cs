using Rubberduck.Parsing.VBA;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface INewDeclarationConflictDetection : IConflictDetectionBase
    {
        /// <summary>
        /// Determines if proposed new module identifier represents a name conflict.
        /// </summary>
        bool NewModuleDeclarationHasConflict(string name, string projectID, IConflictDetectionSessionData sessionData, out string nonConflictName);
    }

    public class NewDeclarationConflictDetection : ConflictDetectionBase, INewDeclarationConflictDetection
    {
        public NewDeclarationConflictDetection(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory) { }

        public override bool HasConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            return !CanResolveToConflictFreeIdentifier(proxy, sessionData);
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
