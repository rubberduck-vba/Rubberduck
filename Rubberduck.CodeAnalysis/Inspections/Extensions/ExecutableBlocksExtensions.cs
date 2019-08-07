using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.Inspections.Extensions
{
    public static class ExecutableBlocksExtensions
    {
        public static bool ContainsExecutableStatements(this BlockContext block)
        {
            return block?.children != null && ContainsExecutableStatements(block.children);
        }

        private static bool ContainsExecutableStatements(System.Collections.Generic.IList<Antlr4.Runtime.Tree.IParseTree> blockChildren)
        {
            foreach (var child in blockChildren)
            {
                if (child is BlockStmtContext blockStmt)
                {
                    var mainBlockStmt = blockStmt.mainBlockStmt();

                    if (mainBlockStmt == null)
                    {
                        continue;   //We have a lone line lable, which is not executable.
                    }

                    // exclude variables and consts because they are not executable statements
                    if (mainBlockStmt.GetChild(0) is VariableStmtContext ||
                        mainBlockStmt.GetChild(0) is ConstStmtContext)
                    {
                        continue;
                    }

                    return true;
                }

                if (child is RemCommentContext ||
                    child is CommentContext ||
                    child is CommentOrAnnotationContext ||
                    child is EndOfStatementContext)
                {
                    continue;
                }

                return true;
            }

            return false;
        }
    }
}
