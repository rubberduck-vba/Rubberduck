using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
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
