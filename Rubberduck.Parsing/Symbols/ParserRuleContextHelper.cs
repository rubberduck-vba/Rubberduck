using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public static IEnumerable<IToken> GetTokens(ParserRuleContext context, CommonTokenStream tokenStream)
        {
            var sourceInterval = context.SourceInterval;
            if (sourceInterval.Equals(Interval.Invalid) || sourceInterval.b < sourceInterval.a)
            {
                return new List<IToken>();
            }
            // Gets the tokens belonging to the context from the token stream. 
           return tokenStream.GetTokens(sourceInterval.a, sourceInterval.b);
        }

        public static T GetChild<T>(RuleContext context)
        {
            if (context == null)
            {
                return default(T);
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index) is T)
                {
                    return (T)child;
                }
            }

            return default(T);
        }

        public static T GetDescendent<T>(RuleContext context)
        {
            if (context == null)
            {
                return default(T);
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index) is T)
                {
                    return (T)child;
                }

                var descendent = GetDescendent<T>(child as RuleContext);
                if (descendent != null)
                {
                    return descendent;
                }
            }

            return default(T);
        }
    }
}
