using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Linq;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing
{
    public static class ParserRuleContextExtensions
    {
        /// <summary>
        ///  Returns a Selection structure containing the context
        /// </summary>
        public static Selection GetSelection(this ParserRuleContext context)
        {
            // if we have an empty module, `Stop` is null
            if (context?.Stop == null) { return Selection.Home; }

            // ANTLR indexes for columns are 0-based, but VBE's are 1-based.
            // ANTLR lines and VBE's lines are both 1-based
            // 1 is the default value that will select all lines. Replace zeroes with ones.
            // See also: https://msdn.microsoft.com/en-us/library/aa443952(v=vs.60).aspx

            return new Selection(context.Start.Line == 0 ? 1 : context.Start.Line,
                                 context.Start.Column + 1,
                                 context.Stop.Line == 0 ? 1 : context.Stop.EndLine(),
                                 context.Stop.EndColumn() + 1);
        }

        /// <summary>
        ///  Gets the tokens belonging to the context from the token stream.
        /// </summary>
        public static IEnumerable<IToken> GetTokens(this ParserRuleContext context, CommonTokenStream tokenStream)
        {
            var sourceInterval = context.SourceInterval;
            if (sourceInterval.Equals(Interval.Invalid) || sourceInterval.b < sourceInterval.a)
            {
                return new List<IToken>();
            }
            return tokenStream.GetTokens(sourceInterval.a, sourceInterval.b);
        }

        /// <summary>
        ///  Gets the original source, without "synthetic" text such as "<EOF>
        /// </summary>
        public static string GetText(this ParserRuleContext context, ICharStream stream)
        {
            // Can be null if the input is empty it seems.
            if (context.Stop == null)
            {
                return string.Empty;
            }
            return stream.GetText(new Interval(context.Start.StartIndex, context.Stop.StopIndex));
        }

        /// <summary>
        /// Returns the first direct child of 'context' that is of the generic Type.
        /// </summary>
        public static TContext GetChild<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            if (context == null)
            {
                return default;
            }

            var results = context.children.Where(child => child is TContext);
            return results.Any() ? (TContext)results.First() : default;
        }

        /// <summary>
        /// Determines if any of the context's ancestors are the generic Type.
        /// </summary>
        public static bool IsDescendentOf<TContext>(this ParserRuleContext context)
        {
            if (context == null)
            {
                return false;
            }

            if (context is TContext)
            {
                return GetAncestor_Recursive<TContext>((ParserRuleContext)context.Parent) != null;
            }
            return GetAncestor_Recursive<TContext>(context) != null;
        }

        /// <summary>
        /// Determines if any of the context's ancestors are equal to the parameter 'ancestor'.
        /// </summary>
        public static bool IsDescendentOf<T>(this ParserRuleContext context, T ancestor) where T : ParserRuleContext
        {
            if (context == null || ancestor == null)
            {
                return false;
            }
            if (context == ancestor)
            {
                return IsDescendentOf_Recursive(context.Parent, ancestor);
            }
            return IsDescendentOf_Recursive(context, ancestor);
        }

        private static bool IsDescendentOf_Recursive(IParseTree context, IParseTree targetParent)
        {
            if (context == null)
            {
                return false;
            }
            if (context == targetParent)
            {
                return true;
            }
            return IsDescendentOf_Recursive(context.Parent, targetParent);
        }

        /// <summary>
        /// Returns the context's first ancestor of the generic Type.
        /// </summary>
        public static TContext GetAncestor<TContext>(this ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }
            if (context is TContext)
            {
                return GetAncestor_Recursive<TContext>((ParserRuleContext)context.Parent);
            }
            return GetAncestor_Recursive<TContext>(context);
        }

        private static TContext GetAncestor_Recursive<TContext>(ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }
            if (context is TContext)
            {
                return (TContext)System.Convert.ChangeType(context, typeof(TContext));
            }
            return GetAncestor_Recursive<TContext>((ParserRuleContext)context.Parent);
        }

        /// <summary>
        /// Returns the context's first descendent of the generic Type.
        /// </summary>
        public static TContext GetDescendent<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            var descendents = GetDescendents<TContext>(context);
            return descendents.Any() ? descendents.First() : null;
        }

        /// <summary>
        /// Returns all the context's descendents of the generic Type.
        /// </summary>
        public static IEnumerable<TContext> GetDescendents<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            var walker = new ParseTreeWalker();
            var listener = new ChildNodeListener<TContext>();
            walker.Walk(listener, context);
            return listener.Matches;
        }

        private class ChildNodeListener<TContext> : VBAParserBaseListener where TContext : ParserRuleContext
        {
            private readonly HashSet<TContext> _matches = new HashSet<TContext>();
            public IEnumerable<TContext> Matches => _matches;

            public override void EnterEveryRule(ParserRuleContext context)
            {
                var match = context as TContext;
                if (match != null)
                {
                    _matches.Add(match);
                }
            }
        }
    }
}