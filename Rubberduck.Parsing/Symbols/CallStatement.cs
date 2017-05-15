using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public static class CallStatement
    {
        public static VBAParser.ArgumentListContext GetArgumentList(VBAParser.CallStmtContext callStmt)
        {
            VBAParser.ArgumentListContext argList = null;
            if (callStmt.CALL() != null && callStmt.lExpression() is VBAParser.IndexExprContext)
            {
                var indexExpr = (VBAParser.IndexExprContext)callStmt.lExpression();
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
