using Rubberduck.RegexAssistant.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.RegexAssistant
{
    public interface IRegularExpression : IDescribable
    {
        Quantifier Quantifier { get; }

        bool TryMatch(string text, out string remaining);
    }

    public class ConcatenatedExpression : IRegularExpression
    {
        private readonly Quantifier quant;
        private readonly IList<IRegularExpression> subexpressions;

        public ConcatenatedExpression(IList<IRegularExpression> subexpressions)
        {
            this.subexpressions = subexpressions;
            this.quant = new Quantifier(""); // these are always exactly once. Quantifying happens through groups
        }

        public string Description
        {
            get
            {
                return string.Join(", ", subexpressions.Select(exp => exp.Description)) + " " + Quantifier.HumanReadable();
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                return quant;
            }
        }

        public bool TryMatch(string text, out string remaining)
        {
            throw new NotImplementedException();
        }
    }

    public class AlternativesExpression : IRegularExpression
    {
        private readonly Quantifier quant;
        private readonly IList<IRegularExpression> subexpressions;

        public AlternativesExpression(IList<IRegularExpression> subexpressions)
        {
            this.subexpressions = subexpressions;
            this.quant = new Quantifier(""); // these are always exactly once. Quantifying happens through groups
        }

        public string Description
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public Quantifier Quantifier
        {
            get
            {
                return quant;
            }
        }

        public bool TryMatch(string text, out string remaining)
        {
            throw new NotImplementedException();
        }
    }

    public class SingleAtomExpression : IRegularExpression
    {
        private readonly Atom atom;
        private readonly Quantifier quant;

        public SingleAtomExpression(Atom atom, Quantifier quant)
        {
            this.atom = atom;
            this.quant = quant;
        }

        public string Description
        {
            get
            {
                return string.Format("{0} {1}.", atom.Description, Quantifier.HumanReadable());
            }
        
        }

        public Quantifier Quantifier
        {
            get
            {
                return quant;
            }
        }

        public bool TryMatch(string text, out string remaining)
        {
            // try to match the atom a given number of times.. 
            throw new NotImplementedException();
        }
    }
        
    public static class RegularExpression
    {
        public static IRegularExpression Parse(string specifier)
        {
            /*
             We basically run a Chain of Responsibility here. At the outermost level, we need to check whether this is an AlternativesExpression.
             If it isn't, we assume it's a ConcatenatedExpression and proceed to create one of these.
             The next step is attempting to parse Atoms. Those are packed into a SingleAtomExpression with their respective Quantifier.

             Note that Atoms can request a Parse of their subexpressions. Prominent example here would be Groups.
             Also note that this here is responsible for separating atoms and Quantifiers. When we matched an Atom we need to try to match a Quantifier and pack them together. 
             If there is no Quantifier following (either because the input is exhausted or there directly is the next atom) then we instead pair with `new Quantifier("")`
             */
            // KISS: Alternatives is when you can split at 
            // grab all indices where we have a pipe
            List<int> pipeIndices = GrabPipeIndices(specifier);
            // and now weed out those inside character classes or groups
            WeedPipeIndices(ref pipeIndices, specifier);
            if (pipeIndices.Count == 0)
            { // assume ConcatenatedExpression when trying to parse all as a single atom fails
                IRegularExpression expression;
                if (TryParseAsAtom(specifier, out expression))
                {
                    return expression;
                }
                else
                {
                    expression = ParseIntoConcatenatedExpression(specifier);
                    return expression;
                }
            }
            else
            {
                return ParseIntoAlternativesExpression(pipeIndices, specifier); 
            }
        }

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

            return null;
        }
    }
}