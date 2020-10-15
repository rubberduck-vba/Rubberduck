using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public interface IExtractInterfaceConflictFinderFactory
    {
        IExtractInterfaceConflictFinder Create(IDeclarationFinderProvider declarationFinderProvider, string projectId);
    }

    public interface IExtractInterfaceConflictFinder
    {
        bool IsConflictingModuleName(string moduleIdentifier);
        string GenerateNoConflictModuleName(string moduleIdentifier);
    }

    public class ExtractInterfaceConflictFinder : IExtractInterfaceConflictFinder
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly string _projectId;

        private static List<DeclarationType> _interfaceIdentifierConstrainingTypes = new List<DeclarationType>()
        {
            DeclarationType.Module,
            DeclarationType.UserDefinedType,
            DeclarationType.Enumeration
        };

        public ExtractInterfaceConflictFinder(IDeclarationFinderProvider declarationFinderProvider, string projectID)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _projectId = projectID;
        }

        public bool IsConflictingModuleName(string moduleIdentifier)
        {
            foreach (var declarationType in _interfaceIdentifierConstrainingTypes)
            {
                var conflictingDeclarations = _declarationFinderProvider.DeclarationFinder.UserDeclarations(declarationType)
                    .Where(d => d.ProjectId == _projectId
                        && IdentifierCheckIsRelevant(d)
                        && d.IdentifierName.Equals(moduleIdentifier, StringComparison.InvariantCultureIgnoreCase));

                if (conflictingDeclarations.Any())
                {
                    return true;
                }
            }

            return false;
        }

        public string GenerateNoConflictModuleName(string moduleIdentifier)
        {
            const int maxAttempts = 100;
            var attempts = 0;
            while (IsConflictingModuleName(moduleIdentifier) && ++attempts < maxAttempts)
            {
                moduleIdentifier = IncrementIdentifier(moduleIdentifier);
            }

            return moduleIdentifier;
        }

        private static bool IdentifierCheckIsRelevant(Declaration declaration)
        {
            return declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedType)
                || declaration.DeclarationType.HasFlag(DeclarationType.Enumeration)
                    ? declaration.Accessibility == Accessibility.Public
                    : true;
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
    }
}
