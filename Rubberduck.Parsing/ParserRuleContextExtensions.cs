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

//TODO: Question - remove?  None of the GetProcedureSelection overloads below are referenced (except for one...in a test)
        //https://github.com/rubberduck-vba/Rubberduck/issues/2164
        //This set of overloads returns the selection for the entire procedure statement body, i.e. Public Function Foo(bar As String) As String
        public static Selection GetProcedureSelection(this VBAParser.FunctionStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.SubStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyGetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyLetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertySetStmtContext context) { return GetProcedureContextSelection(context); }

        private static Selection GetProcedureContextSelection(ParserRuleContext context)
        {
            var endContext1 = context.GetRuleContext<VBAParser.EndOfStatementContext>(0);
            var endContext = context.GetChild<VBAParser.EndOfStatementContext>();
            return new Selection(context.Start.Line == 0 ? 1 : context.Start.Line,
                                 context.Start.Column + 1,
                                 endContext.Start.Line == 0 ? 1 : endContext.Start.Line,
                                 endContext.Start.Column + 1);
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
            if(context is TContext)
            {
                return GetParent_Recursive<TContext>((ParserRuleContext)context.Parent) != null;
            }
            return GetParent_Recursive<TContext>(context) != null;
        }

        /// <summary>
        /// Determines if any of the context's ancestors are equal to the parameter 'parent'.
        /// </summary>
        public static bool IsDescendentOf<T>(this ParserRuleContext context, T parent) where T : ParserRuleContext
        {
            if (context == null || parent == null)
            {
                return false;
            }
            if (context != parent)
            {
                return IsDescendentOf_Recursive(context, parent);
            }
            return false;
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

        //TODO: Question: I would like to rename GetParent<T>(...) function to GetAncestor<T>(...)
        //With a name like GetParent, I would expected the code to simply be: "return context.Parent;"
        //This function is called when the caller is interested in looking 'up' the hierarchy including 
        //and above the immediate parent.

        /// <summary>
        /// Returns the context's first ancestor of the generic Type.
        /// </summary>
        public static TContext GetParent<TContext>(this ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }
            if (context is TContext)
            {
                return GetParent_Recursive<TContext>((ParserRuleContext)context.Parent);
            }
            return GetParent_Recursive<TContext>(context);
        }

        private static TContext GetParent_Recursive<TContext>(ParserRuleContext context)
        {
            if (context == null)
            {
                return default;
            }
            if (context is TContext)
            {
                return (TContext)System.Convert.ChangeType(context, typeof(TContext));
            }
            return GetParent_Recursive<TContext>((ParserRuleContext)context.Parent);
        }

        /// <summary>
        /// Determines if the any of the child contexts child.GetText() equals token .
        /// </summary>
        public static bool HasChildToken(this ParserRuleContext context, string token)
        {
            return context.children.Any(ch => ch.GetText().Equals(token));
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