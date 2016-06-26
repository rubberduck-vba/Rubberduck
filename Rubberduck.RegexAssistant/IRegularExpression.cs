using Rubberduck.RegexAssistant.Extensions;
using Rubberduck.RegexAssistant.i18n;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    internal interface IRegularExpression : IDescribable
    {
        Quantifier Quantifier { get; }
    }

    internal class ConcatenatedExpression : IRegularExpression
    {
        private readonly Quantifier _quantifier;
        internal readonly IList<IRegularExpression> Subexpressions;

        public ConcatenatedExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions;
            _quantifier = new Quantifier(string.Empty); // these are always exactly once. Quantifying happens through groups
        }

        public string Description
        {
            get
            {
                return string.Join(Environment.NewLine, Subexpressions.Select(exp => exp.Description));
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                return _quantifier;
            }
        }
    }

    internal class AlternativesExpression : IRegularExpression
    {
        private readonly Quantifier _quantifier;
        internal readonly IList<IRegularExpression> Subexpressions;

        public AlternativesExpression(IList<IRegularExpression> subexpressions)
        {
            Subexpressions = subexpressions;
            _quantifier = new Quantifier(string.Empty); // these are always exactly once. Quantifying happens through groups
        }

        public string Description
        {
            get
            {
                return AssistantResources.ExpressionDescription_AlternativesExpression + Environment.NewLine + string.Join(Environment.NewLine, Subexpressions.Select(exp => exp.Description));
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                return _quantifier;
            }
        }
    }

    internal class SingleAtomExpression : IRegularExpression
    {
        public readonly IAtom Atom;
        private readonly Quantifier _quantifier;

        public SingleAtomExpression(IAtom atom, Quantifier quantifier)
        {
            Atom = atom;
            _quantifier = quantifier;
        }

        public string Description
        {
            get
            {
                return string.Format("{0} {1}.", Atom.Description, Quantifier.HumanReadable());
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                return _quantifier;
            }
        }

        public override bool Equals(object obj)
        {
            if (obj is SingleAtomExpression)
            {
                SingleAtomExpression other = obj as SingleAtomExpression;
                return other.Atom.Equals(Atom) && other.Quantifier.Equals(Quantifier);
            }
            return false;
        }
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
        /// If there is no Quantifier following (either because the input is exhausted or there directly is the next atom) then we instead pair with `new Quantifier("")` 
        /// </summary>
        /// <param name="specifier">The full Regular Expression specifier to Parse</param>
        /// <returns>An IRegularExpression that encompasses the complete given specifier</returns>
        public static IRegularExpression Parse(string specifier)
        {
            IRegularExpression expression;
            // ByRef requires us to hack around here, because TryParseAsAtom doesn't fail when it doesn't consume the specifier anymore
            string specifierCopy = specifier;
            if (TryParseAsAtom(ref specifierCopy, out expression) && specifierCopy.Length == 0)
            {
                return expression;
            }
            List<IRegularExpression> subexpressions = new List<IRegularExpression>();
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
            List<IRegularExpression> subexpressions = new List<IRegularExpression>();
            string currentSpecifier = specifier;
            while (currentSpecifier.Length > 0)
            {
                IRegularExpression expression;
                // we actually have an AlternativesExpression, return the current status to Parse after updating the specifier
                if (currentSpecifier[0].Equals('|'))
                {
                    specifier = currentSpecifier.Substring(1); // skip leading |
                    return new ConcatenatedExpression(subexpressions);
                }
                if (TryParseAsAtom(ref currentSpecifier, out expression))
                {
                    subexpressions.Add(expression);
                }
            }
            specifier = ""; // we've exhausted the specifier, tell Parse about it to prevent infinite looping
            return new ConcatenatedExpression(subexpressions);
        }

        private static readonly Regex groupWithQuantifier = new Regex("^" + Group.Pattern + Quantifier.Pattern + "?");
        private static readonly Regex characterClassWithQuantifier = new Regex("^" + CharacterClass.Pattern + Quantifier.Pattern + "?");
        private static readonly Regex literalWithQuantifier = new Regex("^" + Literal.Pattern + Quantifier.Pattern + "?");
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
        internal static bool TryParseAsAtom(ref string specifier, out IRegularExpression expression)
        {
            Match m = groupWithQuantifier.Match(specifier);
            if (m.Success)
            {
                string atom = m.Groups["expression"].Value;
                string quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new Group("("+atom+")"), new Quantifier(quantifier));
                return true;
            }
            m = characterClassWithQuantifier.Match(specifier);
            if (m.Success)
            {
                string atom = m.Groups["expression"].Value;
                string quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + 2 + quantifier.Length);
                expression = new SingleAtomExpression(new CharacterClass("["+atom+"]"), new Quantifier(quantifier));
                return true;
            }
            m = literalWithQuantifier.Match(specifier);
            if (m.Success)
            {
                string atom = m.Groups["expression"].Value;
                string quantifier = m.Groups["quantifier"].Value;
                specifier = specifier.Substring(atom.Length + quantifier.Length);
                expression = new SingleAtomExpression(new Literal(atom), new Quantifier(quantifier));
                return true;
            }
            expression = null;
            return false;
        }

        /// <summary>
        /// Makes the given specifier with the given pipeIndices into an AlternativesExpression 
        /// </summary>
        /// <param name="pipeIndices">The indices of Alternative-indicating pipes on the current expression level</param>
        /// <param name="specifier">The specifier to split into subexpressions</param>
        /// <returns>An AlternativesExpression consisting of the split alternatives in the specifier, in order of encounter</returns>
        private static IRegularExpression ParseIntoAlternativesExpression(List<int> pipeIndices, string specifier)
        {
            List<IRegularExpression> expressions = new List<IRegularExpression>();
            string currentRemainder = specifier;
            for (int i = pipeIndices.Count - 1; i > 0; i--)
            {
                expressions.Add(Parse(currentRemainder.Substring(pipeIndices[i] + 1)));
                currentRemainder = currentRemainder.Substring(0, pipeIndices[i] - 1);
            }
            expressions.Reverse(); // because we built them from the back
            return new AlternativesExpression(expressions);
        }


        /// <summary>
        /// Finds all Pipes in the given specifier that are not escaped
        /// </summary>
        /// <param name="specifier">the regex specifier to search for unescaped pipes</param>
        /// <returns>A list populated with the indices of all pipes</returns>
        private static List<int> GrabPipeIndices(string specifier)
        {
            // FIXME: Check assumptions: 
            // - | is never encountered at index 0
            // - | is never preceded by \\

            if (!specifier.Contains("|")) {
                return new List<int>();
            }
            int currentIndex = 0;
            List<int> result = new List<int>();
            while (true)
            {
                currentIndex = specifier.IndexOf("|", currentIndex);
                if(currentIndex == -1)
                {
                    break;
                }
                // ignore escaped literals
                // FIXME this is still a little too naive, since we could actually have something like "\\\|", which means that | is escaped again, but it should suffice for now
                if (currentIndex == 0 || !specifier.Substring(currentIndex - 1, 2).Equals("\\|")
                    || (currentIndex > 1 && specifier.Substring(currentIndex -2, 2).Equals("\\\\")))
                {
                    result.Add(currentIndex);
                }
            }
            return result;
        }

        /// <summary>
        /// Weeds out pipe indices that do not signify alternatives at the current "top level" from the given String.
        /// </summary>
        /// <param name="pipeIndices">indices of unescaped pipes in the given specifier</param>
        /// <param name="specifier">the regex specifier under scrutiny</param>
        internal static void WeedPipeIndices(ref List<int> pipeIndices, string specifier)
        {
            if (pipeIndices.Count == 0)
            {
                return;
            }
            foreach (int pipeIndex in pipeIndices)
            {
                // must not be between () or [] braces, else we just weed it out
                
            }
        }
    }
}