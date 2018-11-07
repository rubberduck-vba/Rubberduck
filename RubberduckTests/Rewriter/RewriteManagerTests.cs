using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class RewriteManagerTests
    {
        [Test]
        [Category("Rewriter")]
        public void ReturnsValidCodePaneSessions()
        {
            var rewritingManager = RewritingManager(out _);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            Assert.IsFalse(codePaneSession.IsInvalidated);
        }


        [Test]
        [Category("Rewriter")]
        public void ReturnsValidAttributesSessions()
        {
            var rewritingManager = RewritingManager(out _);
            var attributesSession = rewritingManager.CheckOutAttributesSession();
            Assert.IsFalse(attributesSession.IsInvalidated);
        }


        [Test]
        [Category("Rewriter")]
        public void InvalidateAllSessionsCallsInvalidateOnAllActiveSessions()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();
            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            rewritingManager.InvalidateAllSessions();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions().Concat(mockFactory.RequestedAttributesSessions()))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Once);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromAnActiveCodePaneSessionCallsInvalidateOnAllActiveSessions()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();
            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            codePaneSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions().Concat(mockFactory.RequestedAttributesSessions()))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Once);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromAnActiveAttributesSessionCallsInvalidateOnAllActiveSessions()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            rewritingManager.CheckOutCodePaneSession();
            var attributesSession =  rewritingManager.CheckOutAttributesSession();
            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            attributesSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions().Concat(mockFactory.RequestedAttributesSessions()))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Once);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromANoLongerActiveCodePaneSessionDoesNotCallInvalidateOnAnyActiveSession_InactiveDueToRewrite()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            var attributesSession = rewritingManager.CheckOutAttributesSession();

            rewritingManager.InvalidateAllSessions();

            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            codePaneSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions()
                .Concat(mockFactory.RequestedAttributesSessions())
                .Where(session => session.Object != codePaneSession && session.Object != attributesSession))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Never);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromANoLongerActiveAttributesSessionDoesNotCallInvalidateOnAnyActiveSession_InactiveDueToRewrite()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            var attributesSession = rewritingManager.CheckOutAttributesSession();

            rewritingManager.InvalidateAllSessions();

            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            codePaneSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions()
                .Concat(mockFactory.RequestedAttributesSessions())
                .Where(session => session.Object != codePaneSession && session.Object != attributesSession))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Never);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromANoLongerActiveCodePaneSessionDoesNotCallInvalidateOnAnyActiveSession_InactiveDueToInvalidateAll()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            var attributesSession = rewritingManager.CheckOutAttributesSession();

            attributesSession.TryRewrite();

            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            codePaneSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions()
                .Concat(mockFactory.RequestedAttributesSessions())
                .Where(session => session.Object != codePaneSession && session.Object != attributesSession))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Never);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void CallingTheRewritingAllowedCallbackFromANoLongerActiveAttributesSessionDoesNotCallInvalidateOnAnyActiveSession_InactiveDueToInvalidateAll()
        {
            var rewritingManager = RewritingManager(out var mockFactory);
            var codePaneSession = rewritingManager.CheckOutCodePaneSession();
            var attributesSession = rewritingManager.CheckOutAttributesSession();

            codePaneSession.TryRewrite();

            rewritingManager.CheckOutCodePaneSession();
            rewritingManager.CheckOutAttributesSession();

            attributesSession.TryRewrite();

            foreach (var mockSession in mockFactory.RequestedCodePaneSessions()
                .Concat(mockFactory.RequestedAttributesSessions())
                .Where( session => session.Object != codePaneSession && session.Object != attributesSession))
            {
                mockSession.Verify(m => m.Invalidate(), Times.Never);
            }
        }

        private IRewritingManager RewritingManager(out MockRewriteSessionFactory mockFactory)
        {
            mockFactory = new MockRewriteSessionFactory();
            return new RewritingManager(mockFactory);
        }
    }

    public class MockRewriteSessionFactory : IRewriteSessionFactory
    {
        private readonly List<Mock<IRewriteSession>> _requestedCodePaneSessions = new List<Mock<IRewriteSession>>();
        private readonly List<Mock<IRewriteSession>> _requestedAttributesSessions = new List<Mock<IRewriteSession>>();

        public IEnumerable<Mock<IRewriteSession>> RequestedCodePaneSessions()
        {
            return _requestedCodePaneSessions;
        }

        public IEnumerable<Mock<IRewriteSession>> RequestedAttributesSessions()
        {
            return _requestedAttributesSessions;
        }

        public IRewriteSession CodePaneSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            var mockSession = MockSession(rewritingAllowed, CodeKind.CodePaneCode);
            _requestedCodePaneSessions.Add(mockSession);
            return mockSession.Object;
        }

        private Mock<IRewriteSession> MockSession(Func<IRewriteSession, bool> rewritingAllowed, CodeKind targetCodeKind)
        {
            var mockSession = new Mock<IRewriteSession>();
            mockSession.Setup(m => m.TryRewrite()).Callback(() => rewritingAllowed.Invoke(mockSession.Object));
            var isInvalidated = false;
            mockSession.Setup(m => m.IsInvalidated).Returns(() => isInvalidated);
            mockSession.Setup(m => m.Invalidate()).Callback(() => isInvalidated = true);
            mockSession.Setup(m => m.TargetCodeKind).Returns(targetCodeKind);

            return mockSession;
        }

        public IRewriteSession AttributesSession(Func<IRewriteSession, bool> rewritingAllowed)
        {
            var mockSession = MockSession(rewritingAllowed, CodeKind.AttributesCode);
            _requestedAttributesSessions.Add(mockSession);
            return mockSession.Object;
        }
    }
}