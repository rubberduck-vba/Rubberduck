using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.Common
{
    public interface IConflictDetectionSession
    {
        /// <summary>
        /// Evaluates the target for Conflicts if relocated to another module
        /// </summary>
        /// <param name="target">The declaration proposed for relocation</param>
        /// <param name="destinationModule"></param>
        /// <param name="accessibility">Uses the Accessibility of the target if no parameter value is provided</param>
        /// <param name="IsMutableIdentifier">
        /// a value of 'true' (the default) allows the session to modify the 
        /// proposed declaration's identifier in order to avoid a naming collision.
        /// When set to 'true' the function will always return 'true' unless an exception occurs.
        /// Explicitly setting a value of 'false' would be appropriate for UI related proposals
        /// where the user is supplying the identifier.
        /// </param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposedRelocation(Declaration target, ModuleDeclaration destinationModule, Accessibility? accessibility = null, bool IsMutableIdentifier = true);

        /// <summary>
        /// Evaluates the target for Conflicts if relocated to another module
        /// </summary>
        /// <remarks>
        /// This override accepts a destination module name in order to support
        /// scenarios where the destination module is new or unknown.  
        /// If the name represents an existing ModuleDeclaration then the call is forwarded to
        /// <see cref="TryProposedRelocation(Declaration, ModuleDeclaration, Accessibility?, bool)"/>
        /// </remarks>
        /// <param name="target"></param>
        /// <param name="destinationModuleName"></param>
        /// <param name="accessibility"></param>
        /// <param name="IsMutableIdentifier">
        /// See IsMutableIdentifier <see cref="TryProposedRelocation(Declaration, ModuleDeclaration, Accessibility?, bool)"/>
        /// </param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposedRelocation(Declaration target, string destinationModuleName, Accessibility? accessibility = null, bool IsMutableIdentifier = true);

        /// <summary>
        /// Evaluates the proposed new Identifier for conflicts
        /// </summary>
        /// <param name="target"></param>
        /// <param name="newName"></param>
        /// <param name="IsMutableIdentifier">
        /// See IsMutableIdentifier <see cref="TryProposedRelocation(Declaration, ModuleDeclaration, Accessibility?, bool)"/>
        /// </param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposeRenamePair(Declaration target, string newName, bool IsMutableNewName = true);

        /// <summary>
        /// Evaluates a proposed new Declaration
        /// </summary>
        /// <param name="name"></param>
        /// <param name="declarationType"></param>
        /// <param name="accessibility"></param>
        /// <param name="destination"></param>
        /// <param name="parentDeclaration"></param>
        /// <param name="idKey">
        /// A value that can be used to associate the conflict evaluation 
        /// results with the result of this function. See <see cref="NewDeclarationIdentifiers"/>
        /// </param>
        /// See IsMutableIdentifier <see cref="TryProposedRelocation(Declaration, ModuleDeclaration, Accessibility?, bool)"/>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposeNewDeclaration(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration destination, Declaration parentDeclaration, out int idKey, bool isMutableIdentifier = true);

        /// <summary>
        /// Evaluates a proposed new Module declaration
        /// </summary>
        /// <param name="name"></param>
        /// <param name="projectID"></param>
        /// <param name="nonConflictName">
        /// The evaluated conflict-free identifier based on the proposed name.  If the
        /// proposed name does not generate conflicts, nonConflictName == name
        /// </param>
        /// <returns></returns>
        bool NewModuleDeclarationHasConflict(string name, string projectID, out string nonConflictName);

        /// <summary>
        /// Provides the results of the ConflictDetectionSession evaluations.
        /// It is possible for a single conflict evaluation request to generate more 
        /// than one rename pair. (e.g. Moving an enumeration may result in multiple enumMember conflicts).
        /// </summary>
        IEnumerable<(Declaration target, string newName)> ConflictFreeRenamePairs { get; }

        /// <summary>
        /// Provides a collection of non-conflict identifierresults for evaluated new Declarations.
        /// The caller is responsible for associating the keyID with the proposed new
        /// declaration conflict evaluation. See <see cref="TryProposeNewDeclaration"/>
        /// </summary>
        IEnumerable<(int keyID, string newName)> NewDeclarationIdentifiers { get; }
    }

    /// <summary>
    /// ConflictDetectionSession provides analysis as to whether or not a proposed change (new name, new location, or new Declaration)
    /// results in a name conflict.  Based on the value of flag IsMutableIdentifier (default is 'true'), the evaluation determines and associates
    /// a conflict-free identifier with the proposed target.  The default algorithm for generating the conflict-free name is to based upon
    /// simply incrementing the proposed name. (e.g. "myValue" => "myValue1", "anotherValue6" => "anotherValue7")
    /// If IsMutableIdentifier is 'false' and a conflict is found, the target declaration will not be present
    /// in the ConflictDetectionSession results provided by <see cref="ConflictFreeRenamePairs"/>.
    /// </summary>
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
