using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public interface IValue
    {
        ValueType ValueType { get; }
        bool AsBool { get; }
        byte AsByte { get; }
        decimal AsDecimal { get; }
        DateTime AsDate { get; }
        string AsString { get; }
    }
}
