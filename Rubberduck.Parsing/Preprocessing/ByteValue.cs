using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ByteValue : IValue
    {
        public ByteValue(byte value)
        {
            AsByte = value;
        }

        public ValueType ValueType => ValueType.Byte;

        public bool AsBool => AsByte != 0;

        public byte AsByte { get; }

        public DateTime AsDate => DateTime.FromOADate(Convert.ToDouble(AsByte));

        public decimal AsDecimal => Convert.ToDecimal(AsByte);

        public string AsString => AsByte.ToString();

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString() => AsByte.ToString();
    }
}
