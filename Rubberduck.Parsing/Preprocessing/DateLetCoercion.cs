using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class DateLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            return new DecimalLetCoercion().ToBool(ToDecimal(value));
        }

        public byte ToByte(object value)
        {
            return new DecimalLetCoercion().ToByte(ToDecimal(value));
        }

        public DateTime ToDate(object value)
        {
            return (DateTime)value;
        }

        public decimal ToDecimal(object value)
        {
            return (decimal)((DateTime)value).ToOADate();
        }

        public string ToString(object value)
        {
            DateTime date = (DateTime)value;
            if (date.Date == VBADateConstants.EPOCH_START.Date)
            {
                return date.ToLongTimeString();
            }
            return date.ToShortDateString();
        }
    }
}
