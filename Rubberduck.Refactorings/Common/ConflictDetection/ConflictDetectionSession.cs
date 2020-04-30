using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionSession
    {
        IEnumerable<(Declaration target, string newName)> ConflictFreeRenamePairs { get; }
        IEnumerable<(int keyID, string newName)> NewDeclarationIdentifiers { get; }
        bool TryProposedRelocation(Declaration target, ModuleDeclaration destinationModule, Accessibility? accessibility = null, bool IsMutableIdentifier = true);
        bool TryProposedRelocation(Declaration target, string destinationModuleName, Accessibility? accessibility = null, bool IsMutableIdentifier = true);
        bool TryProposeRenamePair(Declaration target, string newName, bool IsMutableNewName = true);
        bool TryProposeNewDeclaration(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration destination, Declaration parentDeclaration, out int idKey, bool isMutableIdentifier = true);
        bool NewModuleDeclarationHasConflict(string name, string projectID, out string nonConflictName);
    }

    public class ConflictDetectionSession : IConflictDetectionSession
    {
        private readonly IRenameConflictDetection _renamingConflictDetection;
        private readonly IRelocateConflictDetection _relocatingConflictDetection;
        private readonly INewDeclarationConflictDetection _newDeclarationConflictDetection;

        private readonly IConflictDetectionSessionData _sessionData;

        public ConflictDetectionSession(IConflictDetectionSessionData sessionData, 
                                            IRelocateConflictDetection relocateConflictDetection, 
                                            IRenameConflictDetection renameConflictDetection, 
                                            INewDeclarationConflictDetection newDeclarationConflictDetection)
        {
            _relocatingConflictDetection = relocateConflictDetection;
            _renamingConflictDetection = renameConflictDetection;
            _newDeclarationConflictDetection = newDeclarationConflictDetection;
            _sessionData = sessionData;
        }

        public bool TryProposedRelocation(Declaration target, ModuleDeclaration destinationModule, Accessibility? accessibility = null, bool IsMutableIdentifier = true)
        {
            var proxy = _sessionData.CreateProxy(target, destinationModule.IdentifierName, accessibility);
            var hasConflict = _relocatingConflictDetection.HasConflictInNewLocation(proxy, _sessionData);
            if (hasConflict && !IsMutableIdentifier)
            {
                _sessionData.RemoveProxy(proxy);
                return false;
            }
            return true;
        }

        public bool TryProposedRelocation(Declaration target, string destinationModuleName, Accessibility? accessibility = null, bool IsMutableIdentifier = true)
        {
            var proxy = _sessionData.CreateProxy(target, destinationModuleName, accessibility);
            var hasConflict = _relocatingConflictDetection.HasConflictInNewLocation(proxy, _sessionData);
            if (hasConflict && !IsMutableIdentifier)
            {
                _sessionData.RemoveProxy(proxy);
                return false;
            }
            return true;
        }

        public bool TryProposeRenamePair(Declaration target, string newName, bool IsMutableNewName = true)
        {
            var proxy = _sessionData.CreateProxy(target);
            proxy.IdentifierName = newName;
            var hasConflict = _renamingConflictDetection.HasRenameConflict(proxy, _sessionData);

            if (hasConflict && !IsMutableNewName)
            {
                _sessionData.RemoveProxy(proxy);
                return false;
            }
            return true;
        }

        public bool TryProposeNewDeclaration(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration destination, Declaration parentDeclaration, out int retrievalKey, bool isMutableIdentifier = true)
        {
            var proxy = _sessionData.CreateProxy(name, declarationType, accessibility, destination, parentDeclaration, out retrievalKey);
            var hasConflict = _newDeclarationConflictDetection.NewDeclarationHasConflict(proxy, _sessionData);
            if (hasConflict && !isMutableIdentifier)
            {
                _sessionData.RemoveProxy(proxy);
                return false;
            }
            return true;
        }

        public bool NewModuleDeclarationHasConflict(string name, string projectID, out string nonConflictName)
        {
            return _newDeclarationConflictDetection.NewModuleDeclarationHasConflict(name, projectID, _sessionData, out nonConflictName);
        }

        public IEnumerable<(int keyID, string newName)> NewDeclarationIdentifiers
        {
            get
            {
                var results = new List<(int keyID, string newName)>();
                foreach (var proxy in _sessionData.ResolvedProxyDeclarations)
                {
                    results.Add((proxy.GetHashCode(), proxy.IdentifierName));
                }
                return results;
            }
        }

        public IEnumerable<(Declaration target, string newName)> ConflictFreeRenamePairs
        {
            get
            {
                var results = new List<(Declaration, string)>();
                foreach (var resolvedProxy in _sessionData.ResolvedProxyDeclarations)
                {
                    if (resolvedProxy.Prototype != null && !AreVBAEquivalent(resolvedProxy.IdentifierName, resolvedProxy.Prototype.IdentifierName))
                    {
                        results.Add((resolvedProxy.Prototype, resolvedProxy.IdentifierName));
                    }
                }
                return results;
            }
        }

        private bool AreVBAEquivalent(string idFirst, string idSecond)
            => idFirst.Equals(idSecond, StringComparison.InvariantCultureIgnoreCase);
    }
}
