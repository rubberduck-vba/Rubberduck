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

        //This set of overloads returns the selection for the entire procedure statement body, i.e. Public Function Foo(bar As String) As String
        public static Selection GetProcedureSelection(this VBAParser.FunctionStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.SubStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyGetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyLetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertySetStmtContext context) { return GetProcedureContextSelection(context); }

        private static Selection GetProcedureContextSelection(ParserRuleContext context)
        {
            var endContext = context.GetRuleContext<VBAParser.EndOfStatementContext>(0);
            return new Selection(context.Start.Line == 0 ? 1 : context.Start.Line,
                                 context.Start.Column + 1,
                                 endContext.Start.Line == 0 ? 1 : endContext.Start.Line,
                                 endContext.Start.Column + 1);
        }

        //public static IEnumerable<TContext> FindChildrenX<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        //{
        //    var walker = new ParseTreeWalker();
        //    var listener = new ChildNodeListener<TContext>();
        //    walker.Walk(listener, context);
        //    return listener.Matches;
        //}

        public static IEnumerable<IToken> GetTokens(this ParserRuleContext context, CommonTokenStream tokenStream)
        {
            var sourceInterval = context.SourceInterval;
            if (sourceInterval.Equals(Interval.Invalid) || sourceInterval.b < sourceInterval.a)
            {
                return new List<IToken>();
            }
            // Gets the tokens belonging to the context from the token stream. 
            return tokenStream.GetTokens(sourceInterval.a, sourceInterval.b);
        }

        public static string GetText(this ParserRuleContext context, ICharStream stream)
        {
            // Can be null if the input is empty it seems.
            if (context.Stop == null)
            {
                return string.Empty;
            }
            // Gets the original source, without "synthetic" text such as "<EOF>".
            return stream.GetText(new Interval(context.Start.StartIndex, context.Stop.StopIndex));
        }

        public static T GetChild<T>(this ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index) is T)
                {
                    return (T)child;
                }
            }

            return default;
        }

        //public static IEnumerable<T> GetChildren<T>(this RuleContext context)
        //{
        //    if (context == null)
        //    {
        //        yield break;
        //    }

        //    for (var index = 0; index < context.ChildCount; index++)
        //    {
        //        var child = context.GetChild(index);
        //        if (child is T)
        //        {
        //            yield return (T)child;
        //        }
        //    }
        //}

        public static T GetParent<T>(this ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }
            if (context is T)
            {
                return (T)System.Convert.ChangeType(context, typeof(T));
            }
            //return GetMyParent<T>(context); //.Parent);
            return GetParent<T>((ParserRuleContext)context.Parent);
        }

        //private static T GetMyParent<T>(ParserRuleContext context)
        //{
        //    if (context == null)
        //    {
        //        return default;
        //    }
        //    if (context is T)
        //    {
        //        return (T)System.Convert.ChangeType(context, typeof(T));
        //    }
        //    return GetMyParent<T>((ParserRuleContext)context.Parent);
        //}

        public static bool HasParent<TContext>(this ParserRuleContext context)
        {
            //return TryGetAncestor(context, out TContext _);
            return GetParent<TContext>(context) != null;
        }

        public static bool HasParent(this ParserRuleContext context, IParseTree parent)
        {
            return HasParent((IParseTree)context, parent);
            //if (TryGetAncestor(context, out IParseTree ancestor))
            //{
            //    if (ancestor == parent)
            //    {
            //        return true;
            //    }
            //}
            //return HasParent(ancestor, parent);
        }

        private static bool HasParent(this IParseTree context, IParseTree targetParent)
        {
            if (context == null)
            {
                return false;
            }
            if (context == targetParent)
            {
                return true;
            }
            return HasParent(context.Parent, targetParent);
        }

        private static bool TryGetAncestor<TContext>(this ParserRuleContext context, out TContext ancestor)
        {
            ancestor = GetParent<TContext>(context);
            return ancestor != null;
        }

        public static TContext FindChild<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            if (context == null)
            {
                return default;
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index) is TContext)
                {
                    return (TContext)child;
                }
            }
            return default;
        }

        //public static bool HasChildToken(this IParseTree context, string token)
        //{
        //    for (var index = 0; index < context.ChildCount; index++)
        //    {
        //        var child = context.GetChild(index);
        //        if (context.GetChild(index).GetText().Equals(token))
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        public static bool HasChildToken(this ParserRuleContext context, string token)
        {
            return context.children.Any(ch => ch.GetText().Equals(token));
            //for (var index = 0; index < context.ChildCount; index++)
            //{
            //    var child = context.GetChild(index);
            //    if (context.GetChild(index).GetText().Equals(token))
            //    {
            //        return true;
            //    }
            //}
            //return false;
        }

        public static T GetDescendent<T>(this ParserRuleContext context)
        {
            return GetDescendent<T>((IParseTree)context);
        }

        private static T GetDescendent<T>(this IParseTree context)
        {
            if (context == null)
            {
                return default;
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index) is T)
                {
                    return (T)child;
                }

                var descendent = child.GetDescendent<T>();
                if (descendent != null)
                {
                    return descendent;
                }
            }

            return default;
        }

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