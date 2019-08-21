using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.RegexAssistant.Expressions
{
    public class SingleAtomExpression : IRegularExpression
    {
        public readonly IAtom Atom;

        public SingleAtomExpression(IAtom atom)
        {
            Atom = atom ?? throw new ArgumentNullException();
        }

        public string Description(bool spellOutWhitespace) => $"{Atom.Description(spellOutWhitespace)} {Atom.Quantifier.HumanReadable()}.";

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>(Enumerable.Empty<IRegularExpression>());

        public override string ToString() => $"Atom: {Atom}";
        public override bool Equals(object obj) => obj is SingleAtomExpression other && other.Atom.Equals(Atom);
        public override int GetHashCode() => Atom.GetHashCode();
    }
}