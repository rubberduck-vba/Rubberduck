using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

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

        public static IEnumerable<TContext> FindChildren<TContext>(this ParserRuleContext context) where TContext : ParserRuleContext
        {
            var walker = new ParseTreeWalker();
            var listener = new ChildNodeListener<TContext>();
            walker.Walk(listener, context);
            return listener.Matches;
        }

        public static IEnumerable<T> GetChildren<T>(this RuleContext context)
        {
            if (context == null)
            {
                yield break;
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (child is T)
                {
                    yield return (T)child;
                }
            }
        }

        public static bool HasParent(this RuleContext context, RuleContext parent)
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

        public static bool HasChildToken(this IParseTree context, string token)
        {
            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                if (context.GetChild(index).GetText().Equals(token))
                {
                    return true;
                }
            }
            return false;
        }

        public static T GetDescendent<T>(this IParseTree context)
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

        public static IEnumerable<IParseTree> GetDescendents(this IParseTree context)
        {
            if (context == null)
            {
                yield break;
            }

            for (var index = 0; index < context.ChildCount; index++)
            {
                var child = context.GetChild(index);
                yield return child;

                foreach (var node in child.GetDescendents())
                {
                    yield return node;
                }
            }
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
