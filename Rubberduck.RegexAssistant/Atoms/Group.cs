using Rubberduck.RegexAssistant.i18n;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.RegexAssistant.Atoms
{
    public class Group : IAtom
    {
        public Group(IRegularExpression expression, string specifier, Quantifier quantifier) {
            if (expression == null || quantifier == null)
            {
                throw new ArgumentNullException();
            }

            Quantifier = quantifier;
            Subexpression = expression;
            Specifier = specifier;
        }

        public IRegularExpression Subexpression { get; }

        public Quantifier Quantifier { get; }

        public string Specifier { get; }

        public string Description(bool spellOutWhitespace) => string.Format(AssistantResources.AtomDescription_Group, 
            spellOutWhitespace && WhitespaceToString.IsFullySpellingOutApplicable(Specifier, out var spelledOutWhiteSpace)
                ? spelledOutWhiteSpace
                : Specifier);


        public override string ToString() => Specifier;
        public override bool Equals(object obj)
        {
            return obj is Group other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }
        public override int GetHashCode() => HashCode.Compute(Specifier, Quantifier);
    }
}
