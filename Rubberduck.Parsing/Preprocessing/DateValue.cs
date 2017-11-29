using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class DateValue : IValue
    {
        public DateValue(DateTime value)
        {
            AsDate = value;
        }

        public ValueType ValueType => ValueType.Date;

        public bool AsBool => new DecimalValue(AsDecimal).AsBool;

        public byte AsByte => new DecimalValue(AsDecimal).AsByte;

        public DateTime AsDate { get; }

        public decimal AsDecimal => (decimal)(AsDate).ToOADate();

        public string AsString => AsDate.Date == VBADateConstants.EPOCH_START.Date
            ? AsDate.ToLongTimeString()
            : AsDate.ToShortDateString();

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString() => AsDate.ToString(CultureInfo.InvariantCulture);
    }
}
