using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class DecimalValue : IValue
    {
        public DecimalValue(decimal value)
        {
            AsDecimal = value;
        }

        public ValueType ValueType => ValueType.Decimal;

        public bool AsBool => AsDecimal != 0;

        public byte AsByte => Convert.ToByte(AsDecimal);

        public DateTime AsDate => DateTime.FromOADate(Convert.ToDouble(AsDecimal));

        public decimal AsDecimal { get; }

        public string AsString => AsDecimal.ToString(CultureInfo.InvariantCulture);

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString() => AsDecimal.ToString(CultureInfo.InvariantCulture);
    }
}
