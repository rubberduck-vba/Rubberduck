using Rubberduck.RegexAssistant.i18n;
using System;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant.Atoms
{
    public class Group : IAtom
    {
        private static readonly string Pattern = @"((?<Open>(?<!\\)\()|[^()]*|(?<=\\)[)(]|(?<expression-Open>(?<!\\)\)))*(?(Open)(?!))";
        private static readonly Regex Matcher = new Regex($"^{Pattern}$", RegexOptions.Compiled);

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

        public override bool Equals(object obj)
        {
            return obj is Group other
                && other.Quantifier.Equals(Quantifier)
                && other.Specifier.Equals(Specifier);
        }

        public override int GetHashCode()
        {
            return Specifier.GetHashCode();
        }
    }
}
