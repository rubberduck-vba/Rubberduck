using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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

        public string Description => string.Format(AssistantResources.AtomDescription_Group, Specifier);


        public override string ToString() => Specifier;
        public override bool Equals(object obj)
        {
            return obj is Group other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }

        public override int GetHashCode()
        {
            var hashCode = -613684608;
            hashCode = hashCode * -1521134295 + EqualityComparer<Quantifier>.Default.GetHashCode(Quantifier);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Specifier);
            return hashCode;
        }
    }
}
