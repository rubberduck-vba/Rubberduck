using System;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Rewriter
{
    public class RewritingManager : IRewritingManager
    {
        private readonly HashSet<IRewriteSession> _activeCodePaneSessions = new HashSet<IRewriteSession>();
        private readonly HashSet<IRewriteSession> _activeAttributesSessions = new HashSet<IRewriteSession>();

        private readonly IRewriteSessionFactory _sessionFactory;

        private readonly object _invalidationLockObject = new object();

        public RewritingManager(IRewriteSessionFactory sessionFactory)
        {
            _sessionFactory = sessionFactory;
        }


        public IRewriteSession CheckOutCodePaneSession()
        {
            var newSession = _sessionFactory.CodePaneSession(TryAllowExclusiveRewrite);
            lock (_invalidationLockObject)
            {
                _activeCodePaneSessions.Add(newSession);
            }

            return newSession;
        }

        public IRewriteSession CheckOutAttributesSession()
        {
            var newSession = _sessionFactory.AttributesSession(TryAllowExclusiveRewrite);
            lock (_invalidationLockObject)
            {
                _activeAttributesSessions.Add(newSession);
            }

            return newSession;
        }

        private bool TryAllowExclusiveRewrite(IRewriteSession rewriteSession)
        {
            lock (_invalidationLockObject)
            {
                if (!IsCurrentlyActive(rewriteSession))
                {
                    return false;
                }

                InvalidateAllSessionsInternal();
                return true;
            }
        }

        private bool IsCurrentlyActive(IRewriteSession rewriteSession)
        {
            switch (rewriteSession.TargetCodeKind)
            {
                case CodeKind.CodePaneCode:
                    return _activeCodePaneSessions.Contains(rewriteSession);
                case CodeKind.AttributesCode:
                    return _activeAttributesSessions.Contains(rewriteSession);
                default:
                    throw new NotSupportedException(nameof(rewriteSession));
            }
        }

        public void InvalidateAllSessions()
        {
            lock (_invalidationLockObject)
            {
                InvalidateAllSessionsInternal();
            }
        }

        private void InvalidateAllSessionsInternal()
        {
            foreach (var rewriteSession in _activeCodePaneSessions)
            {
                rewriteSession.Invalidate();
            }
            _activeCodePaneSessions.Clear();

            foreach (var rewriteSession in _activeAttributesSessions)
            {
                rewriteSession.Invalidate();
            }
            _activeAttributesSessions.Clear();
        }
    }
}