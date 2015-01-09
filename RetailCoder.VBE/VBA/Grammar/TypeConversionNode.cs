using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Grammar
{
    [ComVisible(false)]
    public class TypeConversionNode : SyntaxTreeNode
    {
        public TypeConversionNode(Instruction instruction, string scope, Match match = null, IEnumerable<SyntaxTreeNode> childNodes = null) 
            : base(instruction, scope, match)
        {
        }

        private static readonly IDictionary<string, string> Types = new Dictionary<string, string>
        {
            {ReservedKeywords.CBool, ReservedKeywords.Boolean},
            {ReservedKeywords.CByte, ReservedKeywords.Byte},
            {ReservedKeywords.CCur, ReservedKeywords.Currency},
            {ReservedKeywords.CDate, ReservedKeywords.Date},
            {ReservedKeywords.CDbl, ReservedKeywords.Double},
            {ReservedKeywords.CInt, ReservedKeywords.Integer},
            {ReservedKeywords.CLng, ReservedKeywords.Long},
            {ReservedKeywords.CLngLng, ReservedKeywords.LongLong},
            {ReservedKeywords.CLngPtr, ReservedKeywords.Long},
            {ReservedKeywords.CSng, ReservedKeywords.Single},
            {ReservedKeywords.CStr, ReservedKeywords.String},
            {ReservedKeywords.CVar, ReservedKeywords.Variant}
        };

        public string ResultType { get { return Types[RegexMatch.Groups["keyword"].Value]; } }
        public Expression Expression { get { return new Expression(RegexMatch.Groups["expression"].Value); } }
    }
}