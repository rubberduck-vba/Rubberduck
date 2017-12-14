using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public interface IRegularExpression : IDescribable
    {
        IList<IRegularExpression> Subexpressions { get; }
    }

    public class ConcatenatedExpression : IRegularExpression
    {
        public ConcatenatedExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => AssistantResources.ExpressionDescription_ConcatenatedExpression;

        public IList<IRegularExpression> Subexpressions { get; }
    }

    public class AlternativesExpression : IRegularExpression
    {
        public AlternativesExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions ?? throw new ArgumentNullException();
        }

        public string Description => string.Format(AssistantResources.ExpressionDescription_AlternativesExpression, Subexpressions.Count);

        public IList<IRegularExpression> Subexpressions { get; }
    }

    public class SingleAtomExpression : IRegularExpression
    {
        public readonly IAtom Atom;

        public SingleAtomExpression(IAtom atom)
        {
            Atom = atom ?? throw new ArgumentNullException();
        }

        public string Description => $"{Atom.Description} {Atom.Quantifier.HumanReadable()}.";

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>(Enumerable.Empty<IRegularExpression>());

        public override bool Equals(object obj)
        {
            return obj is SingleAtomExpression other
                && other.Atom.Equals(Atom);
        }

        public override int GetHashCode()
        {
            return Atom.GetHashCode();
        }

    }

    public class ErrorExpression : IRegularExpression
    {
        private readonly string _errorToken;

        public ErrorExpression(string errorToken)
        {
            _errorToken = errorToken ?? throw new ArgumentNullException();
        }

        public string Description => string.Format(AssistantResources.ExpressionDescription_ErrorExpression, _errorToken);

        public IList<IRegularExpression> Subexpressions => new List<IRegularExpression>();
    }

    internal static class RegularExpression
    {

        /// <summary>
        /// We basically run a Chain of Responsibility here. At first we try to parse the whole specifier as one Atom.
        /// If this fails, we assume it's a ConcatenatedExpression and proceed to create one of these.
        /// That works well until we encounter a non-escaped '|' outside of a CharacterClass. Then we know that we actually have an AlternativesExpression.
        /// This means we have to check what we got back and add it to a List of subexpressions to the AlternativesExpression. 
        /// We then proceed to the next alternative (ParseIntoConcatenatedExpression consumes the tokens it uses) and keep adding to our subexpressions.
        /// 
        /// Note that Atoms (or more specifically Groups) can request a Parse of their subexpressions. 
        /// Also note that TryParseAtom is responsible for grabbing an Atom <b>and</b> it's Quantifier.
        /// </summary>
        /// <param name="specifier">The full Regular Expression specifier to Parse</param>
        /// <returns>An IRegularExpression that encompasses the complete given specifier</returns>
        public static IRegularExpression Parse(string specifier)
        {
            if (specifier == null)
            {
                throw new ArgumentNullException();
            }

            // ByRef requires us to hack around here, because TryParseAsAtom doesn't fail when it doesn't consume the specifier anymore
            var specifierCopy = specifier;
            if (TryParseAsAtom(ref specifierCopy, out var expression) && specifierCopy.Length == 0)
            {
                return expression;
            }
            var subexpressions = new List<IRegularExpression>();
            while (specifier.Length != 0)
            {
                expression = ParseIntoConcatenatedExpression(ref specifier);
                // ! actually an AlternativesExpression
                if (specifier.Length != 0 || subexpressions.Count != 0)
                {
                    // flatten hierarchy
                    var parsedSubexpressions = (expression as ConcatenatedExpression).Subexpressions;
                    if (parsedSubexpressions.Count == 1)
                    {
                        expression = parsedSubexpressions[0];
                    }
                    subexpressions.Add(expression);
                }
            }
            return (subexpressions.Count == 0) ? expression : new AlternativesExpression(subexpressions);
        }
        /// <summary>
        /// Successively parses the complete specifer into Atoms and returns a ConcatenatedExpression after the specifier has been exhausted or a single '|' is encountered at the start of the remaining specifier.
        /// Note: this may fail to work if the last encountered token cannot be parsed into an Atom, but the remaining specifier has nonzero lenght
        /// </summary>
        /// <param name="specifier">The specifier to Parse into a concatenated expression</param>
        /// <returns>The ConcatenatedExpression resulting from parsing the given specifier, either completely or up to the first encountered '|'</returns>
        private static IRegularExpression ParseIntoConcatenatedExpression(ref string specifier)
        {
            var subexpressions = new List<IRegularExpression>();
            var currentSpecifier = specifier;
            var oldSpecifierLength = currentSpecifier.Length + 1;
            while (currentSpecifier.Length > 0 && currentSpecifier.Length < oldSpecifierLength)
            {
                oldSpecifierLength = currentSpecifier.Length;
                // we actually have an AlternativesExpression, return the current status to Parse after updating the specifier
                if (currentSpecifier[0].Equals('|'))
                {
                    specifier = currentSpecifier.Substring(1); // skip leading |
                    return new ConcatenatedExpression(subexpressions);
                }
                if (TryParseAsAtom(ref currentSpecifier, out var expression))
                {
                    subexpressions.Add(expression);
                }
                else if (currentSpecifier.Length == oldSpecifierLength)
                {
                    subexpressions.Add(new ErrorExpression(currentSpecifier.Substring(0, 1)));
                    currentSpecifier = currentSpecifier.Substring(1);
                }
            }
            specifier = ""; // we've exhausted the specifier, tell Parse about it to prevent infinite looping
            return new ConcatenatedExpression(subexpressions);
        }

        private static readonly Regex groupWithQuantifier = new Regex($"^{Group.Pattern}{Quantifier.Pattern}?", RegexOptions.Compiled);
        private static readonly Regex characterClassWithQuantifier = new Regex($"^{CharacterClass.Pattern}{Quantifier.Pattern}?", RegexOptions.Compiled);
        private static readonly Regex literalWithQuantifier = new Regex($"^{Literal.Pattern}{Quantifier.Pattern}?", RegexOptions.Compiled);
        /// <summary>
        /// Tries to parse the given specifier into an Atom. For that all categories of Atoms are checked in the following order:
        ///  1. Group
        ///  2. Class
        ///  3. Literal
        /// When it succeeds, the given expression will be assigned a SingleAtomExpression containing the Atom and it's Quantifier.
        /// The parsed atom will be removed from the specifier and the method returns true. To check whether the complete specifier was an Atom, 
        /// one needs to examine the specifier after calling this method. If it was, the specifier is empty after calling.
        /// </summary>
        /// <param name="specifier">The specifier to extract the leading Atom out of. Will be shortened if an Atom was successfully extracted</param>
        /// <param name="expression">The resulting SingleAtomExpression</param>
        /// <returns>True, if an Atom could be extracted, false otherwise</returns>
        // Note: could be rewritten to not consume the specifier and instead return an integer specifying the consumed length of specifier. This would remove the by-ref passed string hack
        // internal for testing
        internal static bool TryParseAsAtom(ref string specifier, out IRegularExpression expression)
        {
            var m = groupWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups["expression"].Value;
                var quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new Group($"({atom})", new Quantifier(quantifier)));
                return true;
            }
            m = characterClassWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups["expression"].Value;
                var quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new CharacterClass($"[{atom}]", new Quantifier(quantifier)));
                return true;
            }
            m = literalWithQuantifier.Match(specifier);
            if (m.Success)
            {
                var atom = m.Groups["expression"].Value;
                var quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + quantifier.Length);
                expression = new SingleAtomExpression(new Literal(atom, new Quantifier(quantifier)));
                return true;
            }
            expression = null;
            return false;
        }
    }
}