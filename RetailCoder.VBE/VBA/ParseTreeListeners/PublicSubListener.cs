using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public class PublicSubListener : VBListenerBase, IExtensionListener<VBParser.SubStmtContext>
    {
        private readonly QualifiedModuleName _qualifiedName;
        private readonly IList<QualifiedContext<VBParser.SubStmtContext>> _members = 
            new List<QualifiedContext<VBParser.SubStmtContext>>();

        public IEnumerable<QualifiedContext<VBParser.SubStmtContext>> Members { get { return _members; } }

        public PublicSubListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public override void EnterSubStmt(VBParser.SubStmtContext context)
        {
            var visibility = context.Visibility();
            if (visibility == null || visibility.PUBLIC() != null)
            {
                _members.Add(new QualifiedContext<VBParser.SubStmtContext>(_qualifiedName, context));
            }
        }
    }
}
