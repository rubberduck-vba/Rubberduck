using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Linq;
using Antlr4.Runtime.Misc;
using System;

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

        /// <summary>
        /// Tries to return the context's first ancestor of the generic Type.
        /// </summary>
        public static bool TryGetAncestor<TContext>(this ParserRuleContext context, out TContext ancestor)
        {
            ancestor = context.GetAncestor<TContext>();
            return ancestor != null;
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
        /// Returns the context's first ancestor containing the token with the specified token index or the context itels if it already contains the token.
        /// </summary>
        public static ParserRuleContext GetAncestorContainingTokenIndex(this ParserRuleContext context, int tokenIndex)
        {
            if (context == null)
            {
                return default;
            }

            if (context.ContainsTokenIndex(tokenIndex))
            {
                return context;
            }

            var parent = context.Parent as ParserRuleContext;

            if (parent == null)
            {
                return default;
            }

            return GetAncestorContainingTokenIndex(parent, tokenIndex);
        }

        /// <summary>
        /// Determines whether the context contains the token with the specified token index.
        /// </summary>
        public static bool ContainsTokenIndex(this ParserRuleContext context, int tokenIndex)
        {
            if (context == null)
            {
                return false;
            }

            return context.Start.TokenIndex <= tokenIndex && tokenIndex <= context.Stop.TokenIndex;
        }

        /// <summary>
        /// Returns the context's first descendent of the generic Type.
        /// </summary>
        public static TContext GetDescendent<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            var descendents = GetDescendents<TContext>(context);
            return descendents.FirstOrDefault();
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

        /// <summary>
        /// Try to get the first child of the generic context type.
        /// </summary>
        public static bool TryGetChildContext<TContext>(this ParserRuleContext ctxt, out TContext opCtxt) where TContext : ParserRuleContext
        {
            opCtxt = ctxt.GetChild<TContext>();
            return opCtxt != null;
        }

        /// <summary>
        /// Determines if the context's module declares or defaults to 
        /// Option Compare Binary 
        /// </summary>
        public static bool IsOptionCompareBinary(this ParserRuleContext context)
        {
            if( !(context is VBAParser.ModuleContext moduleContext))
            {
                moduleContext = context.GetAncestor<VBAParser.ModuleContext>();
                if (moduleContext is null)
                {
                    throw new ArgumentException($"Unable to obtain a VBAParser.ModuleContext reference from 'context'");
                }
            }

            var optionContext = moduleContext.GetDescendent<VBAParser.OptionCompareStmtContext>();
            return (optionContext is null) || !(optionContext.BINARY() is null);
        }

        /// Returns the context's first descendent of the generic type containing the token with the specified token index.
        /// </summary>
        public static TContext GetDescendentContainingTokenIndex<TContext>(this ParserRuleContext context, int tokenIndex) where TContext : ParserRuleContext
        {
            var descendents = GetDescendentsContainingTokenIndex<TContext>(context, tokenIndex);
            return descendents.FirstOrDefault();
        }

        /// <summary>
        /// Returns all the context's descendents of the generic type containing the token with the specified token index.
        /// If there are multiple matches, they are ordered from outermost to innermost context.
        /// </summary>
        public static IEnumerable<TContext> GetDescendentsContainingTokenIndex<TContext>(this ParserRuleContext context, int tokenIndex) where TContext : ParserRuleContext
        {
            if (!context.ContainsTokenIndex(tokenIndex))
            {
                return new List<TContext>();
            }

            var matches = new List<TContext>();
            if (context is TContext match)
            {
                matches.Add(match);
            }

            foreach (var child in context.children)
            {
                if (child is ParserRuleContext childContext && childContext.ContainsTokenIndex(tokenIndex))
                {
                    matches.AddRange(childContext.GetDescendentsContainingTokenIndex<TContext>(tokenIndex));
                    break;  //Only one child can contain the token index.
                }
            }

            return matches;
        }

        /// <summary>
        /// Returns the context containing the token preceding the context provided it is of the specified generic type.
        /// </summary>
        public static bool TryGetPrecedingContext<TContext>(this ParserRuleContext context, out TContext precedingContext) where TContext : ParserRuleContext
        {
            precedingContext = null;
            if (context == null)
            {
                return false;
            }

            var precedingTokenIndex = context.Start.TokenIndex - 1;
            var ancestorContainingPrecedingIndex = context.GetAncestorContainingTokenIndex(precedingTokenIndex);

            if (ancestorContainingPrecedingIndex == null)
            {
                return false;
            }

            precedingContext = ancestorContainingPrecedingIndex.GetDescendentContainingTokenIndex<TContext>(precedingTokenIndex);
            return precedingContext != null;
        }

        /// <summary>
        /// Returns the context containing the token following the context provided it is of the specified generic type.
        /// </summary>
        public static bool TryGetFollowingContext<TContext>(this ParserRuleContext context, out TContext followingContext) where TContext : ParserRuleContext
        {
            followingContext = null;
            if (context == null)
            {
                return false;
            }

            var followingTokenIndex = context.Stop.TokenIndex + 1;
            var ancestorContainingFollowingIndex = context.GetAncestorContainingTokenIndex(followingTokenIndex);

            if (ancestorContainingFollowingIndex == null)
            {
                return false;
            }

            followingContext = ancestorContainingFollowingIndex.GetDescendentContainingTokenIndex<TContext>(followingTokenIndex);
            return followingContext != null;
        }

        private class ChildNodeListener<TContext> : VBAParserBaseListener where TContext : ParserRuleContext
        {
            private readonly HashSet<TContext> _matches = new HashSet<TContext>();
            public IEnumerable<TContext> Matches => _matches;

            public override void EnterEveryRule(ParserRuleContext context)
            {
                if (context is TContext match)
                {
                    _matches.Add(match);
                }
            }
        }
    }
}