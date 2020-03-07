using System.Collections.Generic;
using System.Diagnostics;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class EmptyBlockInspectionBase<TContext> : ParseTreeInspectionBase<TContext>
        where TContext : ParserRuleContext
    {
        protected EmptyBlockInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        protected class EmptyBlockInspectionListenerBase : InspectionListenerBase<TContext>
        {
            public void InspectBlockForExecutableStatements<T>(VBAParser.BlockContext block, T context) where T : TContext
            {
                if (!BlockContainsExecutableStatements(block))
                {
                    SaveContext(context);
                }
            }

            private bool BlockContainsExecutableStatements(VBAParser.BlockContext block)
            {
                return block?.children != null && ContainsExecutableStatements(block.children);
            }

            private bool ContainsExecutableStatements(IList<IParseTree> blockChildren)
            {
                foreach (var child in blockChildren)
                {
                    if (child is VBAParser.BlockStmtContext blockStmt)
                    {
                        var mainBlockStmt = blockStmt.mainBlockStmt();

                        if (mainBlockStmt == null)
                        {
                            continue;   //We have a lone line label, which is not executable.
                        }

                        Debug.Assert(mainBlockStmt.ChildCount == 1);

                        // exclude variables and consts because they are not executable statements
                        if (mainBlockStmt.GetChild(0) is VBAParser.VariableStmtContext ||
                            mainBlockStmt.GetChild(0) is VBAParser.ConstStmtContext)
                        {
                            continue;
                        }

                        return true;
                    }

                    if (child is VBAParser.RemCommentContext ||
                        child is VBAParser.CommentContext ||
                        child is VBAParser.CommentOrAnnotationContext ||
                        child is VBAParser.EndOfStatementContext)
                    {
                        continue;
                    }

                    return true;
                }

                return false;
            }

            public void InspectBlockForExecutableStatements<T>(VBAParser.UnterminatedBlockContext block, T context) where T : TContext
            {
                if (!BlockContainsExecutableStatements(block))
                {
                    SaveContext(context);
                }
            }

            private bool BlockContainsExecutableStatements(VBAParser.UnterminatedBlockContext block)
            {
                return block?.children != null && ContainsExecutableStatements(block.children);
            }
        }
    }
}