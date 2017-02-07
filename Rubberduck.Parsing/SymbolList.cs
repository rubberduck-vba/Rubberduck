using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing
{
    public static class SymbolList
    {
        public static readonly IReadOnlyList<string> ValueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.LongPtr,
            Tokens.Single,
            Tokens.String,
            Tokens.Variant,            
        };

        public static readonly IReadOnlyList<string> BaseTypes = new[]
        {
            Tokens.Boolean.ToUpper(),
            Tokens.Byte.ToUpper(),
            Tokens.Currency.ToUpper(),
            Tokens.Date.ToUpper(),
            Tokens.Decimal.ToUpper(),
            Tokens.Double.ToUpper(),
            Tokens.Integer.ToUpper(),
            Tokens.Long.ToUpper(),
            Tokens.LongLong.ToUpper(),
            Tokens.LongPtr.ToUpper(),
            Tokens.Single.ToUpper(),
            Tokens.String.ToUpper(),
            Tokens.Variant.ToUpper(),
            Tokens.Object.ToUpper(),
            Tokens.Any.ToUpper()
        };

        public static readonly IDictionary<string, string> TypeHintToTypeName = new Dictionary<string, string>
        {
            { "%", Tokens.Integer },
            { "&", Tokens.Long },
            { "@", Tokens.Decimal },
            { "!", Tokens.Single },
            { "#", Tokens.Double },
            { "$", Tokens.String }
        };
    }
}
