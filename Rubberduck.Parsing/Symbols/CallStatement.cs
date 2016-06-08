using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public static class CallStatement
    {
        public static VBAParser.ArgumentListContext GetArgumentList(VBAParser.CallStmtContext callStmt)
        {
            VBAParser.ArgumentListContext argList = null;
            if (callStmt.CALL() != null && callStmt.expression() is VBAParser.LExprContext && ((VBAParser.LExprContext)callStmt.expression()).lExpression() is VBAParser.IndexExprContext)
            {
                var indexExpr = (VBAParser.IndexExprContext)((VBAParser.LExprContext)callStmt.expression()).lExpression();
                argList = indexExpr.argumentList();
            }
            else
            {
                argList = callStmt.argumentList();
            }
            return argList;
        }
    }
}
