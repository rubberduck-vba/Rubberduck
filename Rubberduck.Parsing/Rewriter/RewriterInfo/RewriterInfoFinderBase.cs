using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Rewriter.RewriterInfo
{
    public abstract class RewriterInfoFinderBase : IRewriterInfoFinder
    {
        public abstract RewriterInfo GetRewriterInfo(ParserRuleContext context);
        protected static RewriterInfo GetRewriterInfoForTargetRemovedFromListStmt(IToken targetStartToken, int itemIndex, IReadOnlyList<ParserRuleContext> items)
        {
            var count = items.Count;
            var startIndex = itemIndex < count - 1
                ? targetStartToken.TokenIndex
                : items[itemIndex - 1].Stop.TokenIndex + 1;

            var stopIndex = itemIndex < count - 1
                ? items[itemIndex + 1].Start.TokenIndex - 1
                : items[itemIndex].Stop.TokenIndex;

            return new RewriterInfo(startIndex, stopIndex);
        }

        protected static int FindStopTokenIndexForRemoval(VBAParser.ModuleDeclarationsElementContext element)
        {
            if (!element.TryGetFollowingContext<VBAParser.IndividualNonEOFEndOfStatementContext>(out var followingIndividualEndOfLineStatement))
            {
                return element.Stop.TokenIndex;
            }

            //If the endOfStatement starts with a statement separator, it is safe to simply remove that. 
            if (followingIndividualEndOfLineStatement.COLON() != null)
            {
                return followingIndividualEndOfLineStatement.Stop.TokenIndex;
            }

            //Since there is no statement separator, the individual endOfStatement must contain an endOfLine.
            var endOfLine = followingIndividualEndOfLineStatement.endOfLine();

            //EndOfLines contain preceding comments. So, we cannot remove the line, if there is one.
            if (endOfLine.commentOrAnnotation() != null)
            {
                return endOfLine.commentOrAnnotation().Start.TokenIndex - 1;
            }

            return followingIndividualEndOfLineStatement.Stop.TokenIndex;
        }

        protected static int FindStopTokenIndexForRemoval(VBAParser.MainBlockStmtContext mainBlockStmt)
        {
            return FindStopTokenIndexForRemoval((VBAParser.BlockStmtContext)mainBlockStmt.Parent);
        }

        //This overload differs from the one for module declaration elements because we have to take care that we do not invalidate line labels or line numbers on the next line. 
        protected static int FindStopTokenIndexForRemoval(VBAParser.BlockStmtContext blockStmt)
        {
            if (!blockStmt.TryGetFollowingContext<VBAParser.IndividualNonEOFEndOfStatementContext>(out var followingIndividualEndOfLineStatement))
            {
                return blockStmt.Stop.TokenIndex;
            }

            //If the endOfStatement starts with a statement separator, it is safe to simply remove that. 
            if (followingIndividualEndOfLineStatement.COLON() != null)
            {
                return followingIndividualEndOfLineStatement.Stop.TokenIndex;
            }

            //Since there is no statement separator, the individual endOfStatement must contain an endOfLine.
            var endOfLine = followingIndividualEndOfLineStatement.endOfLine();

            //EndOfLines contain preceding comments. So, we cannot remove the line, if there is one.
            if (endOfLine.commentOrAnnotation() != null)
            {
                return endOfLine.commentOrAnnotation().Start.TokenIndex - 1;
            }

            //There could be a statement label right after the individual endOfStatement. 
            //In that case, removing the endOfStatement would break the code.
            if (followingIndividualEndOfLineStatement.TryGetFollowingContext<VBAParser.StatementLabelDefinitionContext>(out _))
            {
                return blockStmt.Stop.TokenIndex;
            }
            
            return followingIndividualEndOfLineStatement.Stop.TokenIndex;
        }

        private static int FindStopTokenIndex(VBAParser.BlockContext context, int index)
        {
            return context.endOfStatement(index).Stop.TokenIndex;
        }

        private static int FindStopTokenIndex(VBAParser.ModuleDeclarationsContext context, int index)
        {
            return context.endOfStatement(index).Stop.TokenIndex;
        }
    }
}