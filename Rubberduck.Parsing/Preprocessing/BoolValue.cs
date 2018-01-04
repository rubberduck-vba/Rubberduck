using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class BoolValue : IValue
    {
        public BoolValue(bool value)
        {
            AsBool = value;
        }

        public ValueType ValueType => ValueType.Bool;

        public bool AsBool { get; }

        public byte AsByte => AsBool 
            ? byte.MaxValue 
            : byte.MinValue;

        public DateTime AsDate => new DecimalValue(AsDecimal).AsDate;

        public decimal AsDecimal => AsBool 
            ? -1 
            : 0;

        public string AsString => AsBool 
            ? bool.TrueString 
            : bool.FalseString;

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString() => AsBool.ToString();
    }
}
