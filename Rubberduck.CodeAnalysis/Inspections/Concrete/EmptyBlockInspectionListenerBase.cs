using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.Concrete
{
    public class EmptyBlockInspectionListenerBase : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public QualifiedModuleName CurrentModuleName { get; set; }

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public void InspectBlockForExecutableStatements<T>(VBAParser.BlockContext block, T context) where T : ParserRuleContext
        {
            if (!BlockContainsExecutableStatements(block))
            {
                AddResult(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }

        public void AddResult(QualifiedContext<ParserRuleContext> qualifiedContext)
        {
            _contexts.Add(qualifiedContext);
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
                        continue;   //We have a lone line lable, which is not executable.
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

        public void InspectBlockForExecutableStatements<T>(VBAParser.UnterminatedBlockContext block, T context) where T : ParserRuleContext
        {
            if (!BlockContainsExecutableStatements(block))
            {
                AddResult(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }

        private bool BlockContainsExecutableStatements(VBAParser.UnterminatedBlockContext block)
        {
            return block?.children != null && ContainsExecutableStatements(block.children);
        }
    }
}
