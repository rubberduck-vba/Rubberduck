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
                switch (lExpr)
                {
                    case VBAParser.IndexExprContext indexExpr1:
                        argList = indexExpr1.argumentList();
                        break;
                    case VBAParser.WhitespaceIndexExprContext indexExpr2:
                        argList = indexExpr2.argumentList();
                        break;
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
