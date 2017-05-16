using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public static class CallStatement
    {
        public static VBAParser.ArgumentListContext GetArgumentList(VBAParser.CallStmtContext callStmt)
        {
            VBAParser.ArgumentListContext argList = null;
            if (callStmt.CALL() != null)
            {
                var lExpr = callStmt.lExpression();
                if (lExpr is VBAParser.IndexExprContext)
                {
                    var indexExpr = (VBAParser.IndexExprContext)lExpr;
                    argList = indexExpr.argumentList();
                }
                else if(lExpr is VBAParser.WhitespaceIndexExprContext)
                {
                    var indexExpr = (VBAParser.WhitespaceIndexExprContext)lExpr;
                    argList = indexExpr.argumentList();
                }
            }
            else
            {
                argList = callStmt.argumentList();
            }
            return argList;
        }
    }
}
