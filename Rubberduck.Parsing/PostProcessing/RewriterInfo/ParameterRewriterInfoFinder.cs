using System;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PostProcessing.RewriterInfo
{
    public class ParameterRewriterInfoFinder : RewriterInfoFinderBase
    {
        public override RewriterInfo GetRewriterInfo(ParserRuleContext context, Declaration target)
        {
            return GetRewriterInfo(target, context.Parent as VBAParser.ArgListContext);
        }

        private static RewriterInfo GetRewriterInfo(Declaration target, VBAParser.ArgListContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context", @"Context is null. Expecting a VBAParser.ArgListContext instance.");
            }

            var items = context.arg();
            var itemIndex = items.ToList().IndexOf((VBAParser.ArgContext)target.Context);
            var count = items.Count;

            if (count == 1)
            {
                return new RewriterInfo(context.LPAREN().Symbol.TokenIndex + 1, context.RPAREN().Symbol.TokenIndex - 1);
            }

            var isLastParam = items.Last() == target.Context;
            if (!isLastParam)
            {
                var removalStop = -1;
                for (var i = context.children.IndexOf(target.Context); i < context.children.Count; i++)
                {
                    var node = context.children[i];
                    if (node.GetText() == ",")
                    {
                        removalStop = (node as TerminalNodeImpl).Symbol.StopIndex;
                    }
                    else if (node is VBAParser.WhiteSpaceContext)
                    {
                        removalStop = (node as VBAParser.WhiteSpaceContext).Stop.StopIndex;
                    }
                    else
                    {
                        break; 
                    }
                }

                var info = GetRewriterInfoForTargetRemovedFromListStmt(target.Context.Start, itemIndex, context.arg());
                return new RewriterInfo(info.StartTokenIndex, removalStop == -1 ? info.StopTokenIndex : removalStop);
            }
            else
            {
                var removalStart = -1;
                for (var i = context.children.IndexOf(target.Context); i >= 0; i--)
                {
                    var node = context.children[i];
                    if (node.GetText() == ",")
                    {
                        removalStart = (node as TerminalNodeImpl).Symbol.StartIndex;
                    }
                    else if (node is VBAParser.WhiteSpaceContext)
                    {
                        removalStart = (node as VBAParser.WhiteSpaceContext).Stop.StartIndex;
                    }
                    else
                    {
                        break;
                    }
                }

                var info = GetRewriterInfoForTargetRemovedFromListStmt(target.Context.Start, itemIndex, context.arg());
                return new RewriterInfo(info.StartTokenIndex, removalStart == -1 ? info.StopTokenIndex : removalStart);
            }
        }
    }
}