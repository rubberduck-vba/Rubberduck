using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public interface IValue
    {
        ValueType ValueType { get; }
        bool AsBool { get; }
        byte AsByte { get; }
        decimal AsDecimal { get; }
        DateTime AsDate { get; }
        string AsString { get; }
        IEnumerable<IToken> AsTokens {get;}
    }
}
