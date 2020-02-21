using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveableMemberSet
    {
        /// <summary>
        /// Returns the wrapped declaration for all types except Properties
        /// Returns the Let/Set declaration of a property unless only a Get exists 
        /// </summary>
        Declaration Member { get; }
        /// <summary>
        /// All declarations sharing the same IdentifierName.
        /// Typically there is only 1 except for Properties
        /// </summary>
        IReadOnlyList<Declaration> Members { get; }
        /// <summary>
        /// The IdentifierName of the wrapped declaration(s)
        /// </summary>
        string IdentifierName { get; }
        /// <summary>
        /// The identifier name to be used when the declaration is moved.
        /// Typically it is equal to the IdentifierName unless there was 
        /// a name conflict in the Destination module
        /// </summary>
        string MovedIdentifierName { set; get; }
        /// <summary>
        /// Is true if MoveIdentifierName != IdentifierName
        /// </summary>
        bool RetainsOriginalIdentifier { get; }
        /// <summary>
        /// IsSelected flags the declaration set 
        /// </summary>
        bool IsSelected { set; get; }
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
        public MoveableMemberSet(Declaration member)
            :this(new List<Declaration>() { member })
        {}

        public MoveableMemberSet(IEnumerable<Declaration> members)
        {
            _members = new List<Declaration>(members);
            MovedIdentifierName = IdentifierName;
        }

        private List<Declaration> _members;
        public IReadOnlyList<Declaration> Members => _members;

        public Declaration Member
        {
            get
            {
                if (_members.Count == 1)
                {
                    return _members.First();
                }
                var subroutinePropertyTypes = _members.Where(m => m.DeclarationType.Equals(DeclarationType.PropertyLet)
                                || m.DeclarationType.Equals(DeclarationType.PropertySet));

                return subroutinePropertyTypes.First();
            }
        }

        public bool IsSelected { set; get; }

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
