using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.CodeAnalysis.Inspections.Extensions
{
    internal static class ExecutableBlocksExtensions
    {
        /// <summary>
        /// Checks a block of code for executable statments and returns true if are present.
        /// </summary>
        /// <param name="block">The block to inspect</param>
        /// <param name="considerAllocations">Determines wheather Dim or Const statements should be considered executables</param>
        /// <returns></returns>
        public static bool ContainsExecutableStatements(this BlockContext block, bool considerAllocations = false)
        {
            return block?.children != null && ContainsExecutableStatements(block.children, considerAllocations);
        }

        private static bool ContainsExecutableStatements(
            IList<IParseTree> blockChildren,
            bool considerAllocations = false)
        {
            foreach (var child in blockChildren)
            {
                if (child is BlockStmtContext blockStmt)
                {
                    var mainBlockStmt = blockStmt.mainBlockStmt();

                    if (mainBlockStmt == null)
                    {
                        continue;   //We have a lone line label, which is not executable.
                    }

                    // if inspection does not consider allocations as executables,
                    // exclude variables and consts because they are not executable statements
                    if (!considerAllocations && IsConstOrVariable(mainBlockStmt.GetChild(0)))
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

        private static bool IsConstOrVariable(IParseTree block)
        {
            return block is VariableStmtContext || block is ConstStmtContext;
        }
    }
}
