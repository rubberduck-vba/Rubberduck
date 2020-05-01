using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// ConflictDetectionSession interface for evaluating proposed renames, relocations, and new declarations
    /// for IdentifierName conflicts and other ambiguous identifier scenarios.
    /// </summary>
    /// <remarks>
    /// The <c>IsMutableIdentifier</c> flag common to the 'TryProposedXXX' functions determines
    /// the behavior of ConflictDetectionSession logic.
    /// Setting this flag to 'true' enables the session logic to modify identifiers as needed 
    /// to avoid/resolve a naming collision.
    /// When set to 'true' <c>TryProposedXXX</c> will always succeed unless an exception occurs.
    /// Setting this flag to 'false' results in a conflict analysis free of any attempt to 
    /// 'coerce' the target's proposed identifier.  Setting a value of 'false' would be appropriate for UI related callers
    /// where the user is supplying the identifier and would not expect it to be modified by the system.
    /// </remarks>
    public interface IConflictDetectionSession
    {
        /// <summary>
        /// Evaluates the target for Conflicts if relocated to another module
        /// </summary>
        /// <param name="target">The declaration proposed for relocation</param>
        /// <param name="destinationModule"></param>
        /// <param name="accessibility">Uses the Accessibility of the target if no parameter value is provided</param>
        /// <param name="IsMutableIdentifier">See <see cref="IConflictDetectionSession"/></param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposedRelocation(Declaration target, ModuleDeclaration destinationModule, bool IsMutableIdentifier, Accessibility? accessibility = null);

        /// <summary>
        /// Evaluates the target for Conflicts if relocated to another module
        /// </summary>
        /// <remarks>
        /// This override accepts a destination module name (string) in order to support
        /// scenarios where the destination module is new or unknown.  
        /// </remarks>
        /// <param name="target"></param>
        /// <param name="destinationModuleName"></param>
        /// <param name="accessibility"></param>
        /// <param name="IsMutableIdentifier">See <see cref="IConflictDetectionSession"/></param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposedRelocation(Declaration target, string destinationModuleName, bool IsMutableIdentifier, Accessibility? accessibility = null);

        /// <summary>
        /// Evaluates the proposed IdentifierName for conflicts
        /// </summary>
        /// <param name="target"></param>
        /// <param name="newName"></param>
        /// <param name="IsMutableIdentifier">See <see cref="IConflictDetectionSession"/></param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposeRenamePair(Declaration target, string newName, bool IsMutableNewName);

        /// <summary>
        /// Evaluates a proposed new Declaration for conflicts with existing code elements
        /// </summary>
        /// <param name="name"></param>
        /// <param name="declarationType"></param>
        /// <param name="accessibility"></param>
        /// <param name="destination"></param>
        /// <param name="parentDeclaration"></param>
        /// <param name="idKey">
        /// A value that can be used to associate the conflict evaluation 
        /// results with the result of the function. See <see cref="NewDeclarationIdentifiers"/>
        /// </param>
        /// <param name="IsMutableIdentifier">See <see cref="IConflictDetectionSession"/></param>
        /// <returns>'true' if there is no conflict, or the conflict has been resolved internally by the ConflictSession</returns>
        bool TryProposeNewDeclaration(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration destination, Declaration parentDeclaration, bool isMutableIdentifier, out int idKey);

        /// <summary>
        /// Evaluates a proposed new Module declaration for conflicts with existing code elements
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
        /// Provides an enumerable set of (<c>Declaration</c>, <c>string</c>) tuples to be used for renaming/resolving 
        /// detected conflicts.
        /// </summary>
        /// <remarks>
        /// 1. The set of tuples contains only (<c>Declaration</c>, <c>string</c>) pairs that change
        /// the IdentifierName of an existing <c>Declaration</c>
        /// 2. It is possible for a single conflict evaluation request to generate more 
        /// than one rename pair. (e.g. Moving an enumeration may result in multiple enumMember conflicts).
        /// </remarks>
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
    /// a conflict-free identifier with the proposed target. 
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

        public bool TryProposedRelocation(Declaration target, ModuleDeclaration destinationModule, bool IsMutableIdentifier, Accessibility? accessibility = null)
        {
            return TryProposedRelocation(target, destinationModule.IdentifierName, IsMutableIdentifier, accessibility);
        }

        public bool TryProposedRelocation(Declaration target, string destinationModuleName, bool isMutableIdentifier, Accessibility? accessibility = null)
        {
            var proxy = _sessionData.CreateProxy(target, destinationModuleName, accessibility);
            proxy.IsMutableIdentifier = isMutableIdentifier;

            var hasConflict = _relocatingConflictDetection.HasConflict(proxy, _sessionData);

            return isMutableIdentifier ? true : !hasConflict;
        }

        public bool TryProposeRenamePair(Declaration target, string newName, bool isMutableIdentifier)
        {
            var proxy = _sessionData.CreateProxy(target);
            proxy.IdentifierName = newName;
            proxy.IsMutableIdentifier = isMutableIdentifier;

            var hasConflict = _renamingConflictDetection.HasConflict(proxy, _sessionData);

            return isMutableIdentifier ? true : !hasConflict;
        }

        public bool TryProposeNewDeclaration(string name, DeclarationType declarationType, Accessibility accessibility, ModuleDeclaration destination, Declaration parentDeclaration, bool isMutableIdentifier, out int retrievalKey)
        {
            var proxy = _sessionData.CreateProxy(name, declarationType, accessibility, destination, parentDeclaration, out retrievalKey);
            proxy.IsMutableIdentifier = isMutableIdentifier;

            var hasConflict = _newDeclarationConflictDetection.HasConflict(proxy, _sessionData);

            return isMutableIdentifier ? true : !hasConflict;
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
                    results.Add((proxy.KeyID, proxy.IdentifierName));
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
