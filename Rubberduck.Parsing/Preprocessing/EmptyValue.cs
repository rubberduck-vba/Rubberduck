using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class EmptyValue : IValue
    {
        public static readonly EmptyValue Value = new EmptyValue();

        public ValueType ValueType => ValueType.Empty;

        public bool AsBool => false;

        public byte AsByte => 0;

        public DateTime AsDate => new DateTime(1899, 12, 30);

        public decimal AsDecimal => 0;

        public string AsString => string.Empty;

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString()
        {
            return "<Empty>";
        } 
    }
}
