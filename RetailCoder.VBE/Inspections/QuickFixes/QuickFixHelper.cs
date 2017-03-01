using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    public static class QuickFixHelper
    {
        
        public static IReadOnlyList<VBAParser.BlockStmtContext> GetBlockStmtContexts(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return ((VBAParser.SubStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return ((VBAParser.FunctionStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return ((VBAParser.PropertyLetStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return ((VBAParser.PropertyGetStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return ((VBAParser.PropertySetStmtContext)context).block().blockStmt();
            }
            return Enumerable.Empty<VBAParser.BlockStmtContext>().ToArray();
        }
        
        public static IReadOnlyList<VBAParser.ArgContext> GetArgContexts(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return ((VBAParser.SubStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return ((VBAParser.FunctionStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return ((VBAParser.PropertyLetStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return ((VBAParser.PropertyGetStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return ((VBAParser.PropertySetStmtContext)context).argList().arg();
            }
            return Enumerable.Empty<VBAParser.ArgContext>().ToArray();
        }
    }
}
