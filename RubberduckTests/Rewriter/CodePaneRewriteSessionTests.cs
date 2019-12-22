using System;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class CodePaneRewriteSessionTests : RewriteSessionTestBase
    {
        [Test]
        [Category("Rewriter")]
        public void RequestsParseAfterRewriteIfNotInvalidatedAndParsingAllowed()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnParseRequested(It.IsAny<object>()));

            var rewriteSession = RewriteSession(mockParseManager.Object, session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(otherModule);

            rewriteSession.TryRewrite();

            mockParseManager.Verify(m => m.OnParseRequested(It.IsAny<object>()), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRequestsParseAfterRewriteIfInvalidated()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnParseRequested(It.IsAny<object>()));

            var rewriteSession = RewriteSession(mockParseManager.Object, session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
            rewriteSession.CheckOutModuleRewriter(otherModule);

            rewriteSession.TryRewrite();

            mockParseManager.Verify(m => m.OnParseRequested(It.IsAny<object>()), Times.Never);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotRequestsParseAfterRewriteIfNotParsingAllowed()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnParseRequested(It.IsAny<object>()));

            var rewriteSession = RewriteSession(mockParseManager.Object, session => false, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(otherModule);

            rewriteSession.TryRewrite();

            mockParseManager.Verify(m => m.OnParseRequested(It.IsAny<object>()), Times.Never);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotRequestParseIfNoRewritersAreCheckedOut()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnParseRequested(It.IsAny<object>()));
            var rewriteSession = RewriteSession(mockParseManager.Object, session => true, out _);

            rewriteSession.TryRewrite();

            mockParseManager.Verify(m => m.OnParseRequested(It.IsAny<object>()), Times.Never);
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsTrueIfNotInvalidatedAndParsingAllowed()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            var actual = rewriteSession.TryRewrite();
            Assert.IsTrue(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void TargetsCodePaneCode()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            Assert.AreEqual(CodeKind.CodePaneCode, rewriteSession.TargetCodeKind);
        }

        protected override IExecutableRewriteSession RewriteSession(IParseManager parseManager, Func<IRewriteSession, bool> rewritingAllowed, out MockRewriterProvider mockProvider, bool rewritersAreDirty = false, ISelectionRecoverer selectionRecoverer = null)
        {
            if (selectionRecoverer == null)
            {
                selectionRecoverer = new Mock<ISelectionRecoverer>().Object;
            }
            mockProvider = new MockRewriterProvider(rewritersAreDirty);
            return new CodePaneRewriteSession(parseManager, mockProvider, selectionRecoverer, rewritingAllowed);
        }
    }
}