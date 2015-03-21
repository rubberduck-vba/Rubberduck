using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Listeners
{
    public class PublicSubListener : VBABaseListener, IExtensionListener<VBAParser.SubStmtContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBAParser.SubStmtContext>> _members = 
            new List<QualifiedContext<VBAParser.SubStmtContext>>();

        public IEnumerable<QualifiedContext<VBAParser.SubStmtContext>> Members { get { return _members; } }

        public PublicSubListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public override void EnterSubStmt(VBAParser.SubStmtContext context)
        {
            var visibility = context.visibility();
            if (visibility == null || visibility.PUBLIC() != null)
            {
                _members.Add(new QualifiedContext<VBAParser.SubStmtContext>(_qualifiedName, context));
            }
        }
    }
}
