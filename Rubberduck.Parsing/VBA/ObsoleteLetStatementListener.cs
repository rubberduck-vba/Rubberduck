using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.VBA
{
    public class ObsoleteLetStatementListener : VBABaseListener
    {
        private readonly IList<VBAParser.LetStmtContext> _contexts = new List<VBAParser.LetStmtContext>();
        public IEnumerable<VBAParser.LetStmtContext> Contexts { get { return _contexts; } }

        public override void EnterLetStmt(VBAParser.LetStmtContext context)
        {
            if (context.LET() != null)
            {
                _contexts.Add(context);
            }
        }
    }
}