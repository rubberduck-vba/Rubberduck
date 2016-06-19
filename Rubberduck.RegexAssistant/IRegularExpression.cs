using Rubberduck.RegexAssistant.Extensions;
using System;
using System.Collections.Generic;

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

        public ConcatenatedExpression(IList<IRegularExpression> subexpressions, Quantifier quant)
        {
            this.quant = quant;
            this.subexpressions = subexpressions;
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

    public class AlternativesExpression : IRegularExpression
    {
        private readonly Quantifier quant;
        private readonly IList<IRegularExpression> subexpressions;

        public AlternativesExpression(IList<IRegularExpression> subexpressions, Quantifier quant)
        {
            this.subexpressions = subexpressions;
            this.quant = quant;
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

            return null;
        }
    }
}
