using System;
using System.Text.RegularExpressions;

namespace Rubberduck.RegexAssistant
{
    public class Quantifier
    {
        public static readonly string Pattern = @"(?<quantifier>(?<!\\)[\?\*\+]|(?<!\\)\{(\d+)(,\d*)?(?<!\\)\})";
        private static readonly Regex Matcher = new Regex(@"^\{(?<min>\d+)(?<max>,\d*)?\}$");

        public readonly QuantifierKind Kind;
        public readonly int MinimumMatches;
        public readonly int MaximumMatches;

        public Quantifier(string expression)
        {
            if (expression.Length == 0)
            {
                Kind = QuantifierKind.None;
                MaximumMatches = 1;
                MinimumMatches = 1;
            }
            else if (expression.Length > 1)
            {
                Kind = QuantifierKind.Expression;
                Match m = Matcher.Match(expression);
                if (!m.Success)
                {
                    throw new ArgumentException(String.Format("Cannot extract a Quantifier from the expression {1}", expression));
                }
                int minimum;
                // shouldn't ever happen
                if (!int.TryParse(m.Groups["min"].Value, out minimum))
                {
                    throw new ArgumentException("Cannot Parse Quantifier Expression into Range");
                }
                MinimumMatches = minimum;

                string maximumString = m.Groups["max"].Value; // drop the comma
                if (maximumString.Length > 1)
                {
                    int maximum;
                    // shouldn't ever happen
                    if (!int.TryParse(maximumString.Substring(1), out maximum))
                    {
                        throw new ArgumentException("Cannot Parse Quantifier Expression into Range");
                    }
                    MaximumMatches = maximum;
                }
                else if (maximumString.Length == 1) // got a comma, so we're unbounded
                {
                    MaximumMatches = int.MaxValue;
                }
                else // exact match, because no comma
                {
                    MaximumMatches = minimum;
                }
            }
            else
            {
                switch (expression.ToCharArray()[0])
                {
                    case '*':
                        MinimumMatches = 0;
                        MaximumMatches = int.MaxValue;
                        Kind = QuantifierKind.Wildcard;
                        break;
                    case '+':
                        MinimumMatches = 1;
                        MaximumMatches = int.MaxValue;
                        Kind = QuantifierKind.Wildcard;
                        break;
                    case '?':
                        MinimumMatches = 0;
                        MaximumMatches = 1;
                        Kind = QuantifierKind.Wildcard;
                        break;
                    default:
                        throw new ArgumentException("Passed Quantifier String was not an allowed Quantifier");
                }
            }
        }

        public override bool Equals(object obj)
        {
            if (obj is Quantifier)
            {
                var other = obj as Quantifier;
                return other.Kind == Kind && other.MinimumMatches == MinimumMatches && other.MaximumMatches == MaximumMatches;
            }
            return false;
        }

        public override int GetHashCode()
        {
            // FIXME get a proper has function
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("Quantifier[{0}: {1} to {2}", Kind, MinimumMatches, MaximumMatches);
        }
    }

    public enum QuantifierKind
    {
        None, Expression, Wildcard
    }
}
