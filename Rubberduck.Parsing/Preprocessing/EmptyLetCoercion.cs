using System;

namespace Rubberduck.Parsing.Preprocessing
{
    /// <summary>
    /// Defines the neutral element for each type.
    /// </summary>
    public sealed class EmptyLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            return false;
        }

        public byte ToByte(object value)
        {
            return 0;
        }

        public DateTime ToDate(object value)
        {
            return new DateTime(1899, 12, 30);
        }

        public decimal ToDecimal(object value)
        {
            return 0;
        }

        public string ToString(object value)
        {
            return string.Empty;
        }
    }
}
