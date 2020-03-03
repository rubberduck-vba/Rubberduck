using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveableMemberSet
    {
        /// <summary>
        /// Properties: Returns the first Property member in the Members collection  
        /// Non-Properties: Returns the wrapped declaration for all types except Properties
        /// </summary>
        Declaration Member { get; }

        /// <summary>
        /// All declarations sharing the same IdentifierName.
        /// Typically there is only 1 declaration in the list except for Properties
        /// </summary>
        IReadOnlyList<Declaration> Members { get; }

        /// <summary>
        /// The IdentifierName of the wrapped declaration(s)
        /// </summary>
        string IdentifierName { get; }

        /// <summary>
        /// The identifier name to be used when the declaration is moved.
        /// Typically it is equal to the IdentifierName unless there is 
        /// a name conflict in the Destination module
        /// </summary>
        string MovedIdentifierName { set; get; }

        /// <summary>
        /// Returns true if MoveIdentifierName != IdentifierName
        /// </summary>
        bool RetainsOriginalIdentifier { get; }

        /// <summary>
        /// Set to true if the declaration set as a defining element of the move 
        /// </summary>
        bool IsSelected { set; get; }

        /// <summary>
        /// Returns true if the MoveableMemberSet contains the declaration 
        /// </summary>
        bool Contains(Declaration declaration);

        /// <summary>
        /// Set to true if the declaration is referenced by the MoveMember CallTree 
        /// </summary>
        bool IsSupport { set; get; }

        /// <summary>
        /// Set to true if the declaration is referenced exclusively by the MoveMember participants 
        /// </summary>
        bool IsExclusive { set; get; }
        
        ///// <summary>
        ///// Returns references other than those local to the Member body.  e.g, Function return assignments 
        ///// </summary>
        //IEnumerable<IdentifierReference> NonMemberBodyReferences { get; }

        /// <summary>
        /// Set to true if all Members have Private Accessibility 
        /// </summary>
        bool HasPrivateAccessibility { get; }

        /// <summary>
        /// Returns the direct dependency declarations for the MoveableMemberSet 
        /// </summary>
        IReadOnlyCollection<Declaration> DirectDependencies { get; }

        /// <summary>
        /// Returns the a flatted collection of the call-tree dependencies for the MoveableMemberSet 
        /// </summary>
        IReadOnlyCollection<Declaration> FlattenedDependencies { set; get; }
    }

    /// <summary>
    /// MoveableMemberSet is a set of declarations with the same Identifier.
    /// It binds together all Properties of the same identifier as a single 
    /// element to assist with logic and moves.  MemberMoves are based on the
    /// IsSelected property and enforces an 'all or none' selection
    /// rule for property members.
    /// All other declarations are a MoveableMemberSet with a single declaration
    /// </summary>
    public class MoveableMemberSet : IMoveableMemberSet
    {
        //private List<IdentifierReference> _containedReferences;
        public MoveableMemberSet(Declaration member)
            : this(new List<Declaration>() { member })
        { }

        public MoveableMemberSet(IEnumerable<Declaration> members)
        {
            _members = members.ToList();
            MovedIdentifierName = IdentifierName;
        }

        private List<Declaration> _members;
        public IReadOnlyList<Declaration> Members => _members;

        public IEnumerable<IdentifierReference> ContainedReferences {set; get;}

        public IReadOnlyCollection<Declaration> DirectDependencies 
            => ContainedReferences.Select(rf => rf.Declaration).Distinct().ToList();

        public IReadOnlyCollection<Declaration> FlattenedDependencies { set; get; } = new List<Declaration>();

        public Declaration Member => _members.First();

        public bool IsSelected { set; get; }

        public bool IsSupport { set; get; }

        public bool IsExclusive { set; get; }

        public bool HasPrivateAccessibility => Members.All(mm => mm.HasPrivateAccessibility());

        public bool Contains(Declaration declaration) => Members.Contains(declaration);

        public string IdentifierName => _members.First().IdentifierName;

        public string MovedIdentifierName { set; get; }

        public bool RetainsOriginalIdentifier 
            => MovedIdentifierName.IsEquivalentVBAIdentifierTo(IdentifierName);

        public override bool Equals(object obj)
        {
            if (obj is MoveableMemberSet mm)
            {
                return mm.IdentifierName == IdentifierName
                    && mm.MovedIdentifierName == MovedIdentifierName;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return IdentifierName.GetHashCode();
        }

        public override string ToString()
        {
            return IdentifierName;
        }
    }
}
