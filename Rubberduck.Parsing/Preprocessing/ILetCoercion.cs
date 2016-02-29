using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public interface ILetCoercion
    {
        bool ToBool(object value);
        byte ToByte(object value);
        decimal ToDecimal(object value);
        DateTime ToDate(object value);
        string ToString(object value);
    }
}
