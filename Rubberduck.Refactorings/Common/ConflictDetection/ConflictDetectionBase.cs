using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionBase
    {
        Func<string, string> ConflictingNameModifier { set; get; }
    }

    /// <summary>
    /// Base class for all ConflictDetection classes. ConflictDetection classes are stateless.  
    /// They operate on/with the <see cref="ConflictDetectionSessionData"/>.
    /// <seealso cref="RenameConflictDetection"/>
    /// <seealso cref="RelocateConflictDetection"/>
    /// <seealso cref="NewDeclarationConflictDetection"/>
    /// </summary>
    public class ConflictDetectionBase : IConflictDetectionBase
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IConflictFinderFactory _conflictFinderFactory;

        public ConflictDetectionBase(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _conflictFinderFactory = conflictFinderFactory;
        }

        public Func<string, string> ConflictingNameModifier { set; get; } = IncrementIdentifier;

        protected IEnumerable<Declaration> ModuleIdentifierConflicts(string name, string projectID)
        {
            
            var matchingModules = _declarationFinderProvider.DeclarationFinder.AllModules
                .Where(mod => AreVBAEquivalent(mod.ComponentName, name) && mod.ProjectId == projectID)
                .Select(qmn => _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(qmn));

            var matchingUDTs = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.UserDefinedType)
                .Where(udt => AreVBAEquivalent(udt.IdentifierName, name) && udt.Accessibility != Accessibility.Private && udt.ProjectId == projectID);

            var matchingEnums = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Enumeration)
                .Where(en => AreVBAEquivalent(en.IdentifierName, name) && en.Accessibility != Accessibility.Private && en.ProjectId == projectID);

            return matchingModules.Concat(matchingUDTs).Concat(matchingEnums);
        }

        protected bool TryResolveToConflictFreeIdentifier(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            var isConflictFree = false;
            var iterationMax = 100;
            for (var iteration = 0; iteration < iterationMax && !isConflictFree; iteration++)
            {
                if (HasNameConflicts(proxy, sessionData, out var conflicts))
                {
                    foreach (var conflict in conflicts)
                    {
                        conflict.Key.IdentifierName = ConflictingNameModifier(conflict.Key.IdentifierName);
                        sessionData.RegisterResolvedProxyIdentifier(conflict.Key);
                    }
                    continue;
                }
                isConflictFree = true;
            }
            return isConflictFree;
        }

        protected bool HasNameConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            if (!IsPotentialProjectNameConflictType(proxy.DeclarationType)
                && !IdentifierIsUsedElsewhereInProject(proxy, sessionData))
            {
                return false;
            }

            var conflictFinder = _conflictFinderFactory.Create(proxy.DeclarationType);

            return conflictFinder.TryFindConflict(proxy, sessionData, out conflicts);
        }

        public bool IdentifierIsUsedElsewhereInProject(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
            => IdentifierIsUsedElsewhereInProject(proxy.IdentifierName, proxy.ProjectId)
            || sessionData.ResolvedProxyDeclarations.Any(pxy => AreVBAEquivalent(proxy.IdentifierName, pxy.IdentifierName));

        protected bool IsPotentialProjectNameConflictType(DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Enumeration)
                || declarationType.HasFlag(DeclarationType.UserDefinedType)
                || declarationType.HasFlag(DeclarationType.Project);
        }

        private static string IncrementIdentifier(string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }

        protected bool AreVBAEquivalent(string idFirst, string idSecond)
            => idFirst.Equals(idSecond, StringComparison.InvariantCultureIgnoreCase);

        private bool IdentifierIsUsedElsewhereInProject(string identifier, string projectID)
            => _declarationFinderProvider.DeclarationFinder.MatchName(identifier)
                            .Any(matchedName => matchedName.ProjectId == projectID);
    }
}
