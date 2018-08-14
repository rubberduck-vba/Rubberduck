using Antlr4.Runtime;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ComparableDateValue : IValue, IComparable<ComparableDateValue>
    {
        private readonly DateValue _inner;
        private readonly int _hashCode;

        public ComparableDateValue(DateValue dateValue)
        {
            _inner = dateValue;
            _hashCode = dateValue.AsDecimal.GetHashCode();
        }

        public Parsing.PreProcessing.ValueType ValueType => _inner.ValueType;

        public bool AsBool => _inner.AsBool;

        public byte AsByte => _inner.AsByte;

        public decimal AsDecimal => _inner.AsDecimal;

        public DateTime AsDate => _inner.AsDate;

        public string AsString => _inner.AsString;

        public IEnumerable<IToken> AsTokens => _inner.AsTokens;

        public int CompareTo(ComparableDateValue dateValue)
            => _inner.AsDecimal.CompareTo(dateValue._inner.AsDecimal);

        public override int GetHashCode() => _hashCode;

        public override bool Equals(object obj)
        {
            if (obj is ComparableDateValue decorator)
            {
                return decorator.CompareTo(this) == 0;
            }

            if (obj is DateValue dateValue)
            {
                return dateValue.AsDecimal == _inner.AsDecimal;
            }

            return false;
        }

        public override string ToString()
        {
            return _inner.ToString();
        }

        public string AsDateLiteral()
        {
            return AsDateLiteral(ToString());
        }

        private static string AsDateLiteral(string input)
        {
            var prePostPend = "#";
            var result = input.StartsWith(prePostPend) ? input : $"{prePostPend}{input}";
            result = result.EndsWith(prePostPend) ? result : $"{result}{prePostPend}";
            result.Replace(" 00:00:00", "");
            return result;
        }
    }
}
