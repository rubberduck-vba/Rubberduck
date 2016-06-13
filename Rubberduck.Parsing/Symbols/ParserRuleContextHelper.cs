using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using System;

namespace Rubberduck.Parsing.Symbols
{
    public static class ParserRuleContextHelper
    {
        public static bool HasParent<T>(RuleContext context)
        {
            return GetParent<T>(context) != null;
        }

        public static bool HasParent(RuleContext context, RuleContext parent)
        {
            if (context == null)
            {
                return false;
            }
            if (context == parent)
            {
                return true;
            }
            return HasParent(context.Parent, parent);
        }

        public static T GetParent<T>(RuleContext context)
        {
            if (context == null)
            {
                return default(T);
            }
            if (context is T)
            {
                return (T)Convert.ChangeType(context, typeof(T));
            }
            return GetParent<T>(context.Parent);
        }

        public static string GetText(ParserRuleContext context, ICharStream stream)
        {
            // Can be null if the input is empty it seems.
            if (context.Stop == null)
            {
                return string.Empty;
            }
            // Gets the original source, without "synthetic" text such as "<EOF>".
            return stream.GetText(new Interval(context.Start.StartIndex, context.Stop.StopIndex));
        }
    }
}
