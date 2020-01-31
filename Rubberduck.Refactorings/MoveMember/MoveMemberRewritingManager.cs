using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRewritingManager : IRewritingManager
    {
        private IRewritingManager _rewritingManager;
        public IExecutableRewriteSession MoveRewriteSession { get; }

        public MoveMemberRewritingManager(IRewritingManager rewritingManager)
        {
            _rewritingManager = rewritingManager;
            MoveRewriteSession = rewritingManager.CheckOutCodePaneSession();
        }

        public IMoveEndpointRewriter CheckOutTemporaryRewriter(QualifiedModuleName qmn)
        {
            var session = CheckOutCodePaneSession();
            return new MoveMemberEndpointRewriter(session.CheckOutModuleRewriter(qmn));
        }

        public IMoveEndpointRewriter CheckOutEndpointRewriter(QualifiedModuleName qmn)
        {
            var rewriter = MoveRewriteSession.CheckOutModuleRewriter(qmn);
            return new MoveMemberEndpointRewriter(rewriter);
        }

        public IModuleRewriter CheckOutModuleRewriter(QualifiedModuleName qmn)
        {
            return MoveRewriteSession.CheckOutModuleRewriter(qmn);
        }

        public IExecutableRewriteSession CheckOutCodePaneSession()
        {
            return _rewritingManager.CheckOutCodePaneSession();
        }

        public IExecutableRewriteSession CheckOutAttributesSession()
        {
            return _rewritingManager.CheckOutAttributesSession();
        }

        public void InvalidateAllSessions()
        {
            _rewritingManager.InvalidateAllSessions();
        }
    }
}
