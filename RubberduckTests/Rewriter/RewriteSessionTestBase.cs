using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public abstract class RewriteSessionTestBase
    {
        [Test]
        [Category("Rewriter")]
        public void IsNotInvalidatedAtStart()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            Assert.IsFalse(rewriteSession.IsInvalidated);
        }

        [Test]
        [Category("Rewriter")]
        public void IsInvalidatedAfterCallingInvalidate()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            rewriteSession.Invalidate();
            Assert.IsTrue(rewriteSession.IsInvalidated);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRewritingAllowedOnRewriteIfNoRewritersHaveBeenCheckedOut()
        {
            var isCalled = false;
            var rewriteSession = RewriteSession(session =>
            {
                isCalled = true;
                return true;
            }, out _);
            rewriteSession.TryRewrite();
            Assert.IsFalse(isCalled);
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsFalseIfNoRewritersHaveBeenCheckedOut()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            var actual = rewriteSession.TryRewrite();
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void CallsRewritingAllowedOnRewriteIfNotInvalidatedAndRewritersHaveBeenCheckedOut()
        {
            var isCalled = false; 
            var rewriteSession = RewriteSession(session =>
                {
                    isCalled = true;
                    return true;
                }, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();
            Assert.IsTrue(isCalled);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRewritingAllowedOnRewriteIfInvalidated()
        {
            var isCalled = false;
            var rewriteSession = RewriteSession(session =>
            {
                isCalled = true;
                return true;
            }, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.Invalidate();
            rewriteSession.CheckOutModuleRewriter(otherModule);

            rewriteSession.TryRewrite();
            Assert.IsFalse(isCalled);
        }

        [Test]
        [Category("Rewriter")]
        public void CallsRewriteOnAllCheckedOutRewritersIfNotInvalidatedAndParsingAllowed()
        {
            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty,"TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(otherModule);

            var requestedRewriters = mockRewriterProvider.RequestedRewriters();

            rewriteSession.TryRewrite();

            foreach (var (qmn, codeKind, mockRewriter) in requestedRewriters)
            {
                mockRewriter.Verify(m => m.Rewrite(), Times.Once);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsFalseIfInvalidated()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.Invalidate();
            var actual = rewriteSession.TryRewrite();
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRewriteOnAnyCheckedOutRewriterIfInvalidated()
        {
            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.Invalidate();
            rewriteSession.CheckOutModuleRewriter(otherModule);

            var requestedRewriters = mockRewriterProvider.RequestedRewriters();

            rewriteSession.TryRewrite();

            foreach (var (qmn, codeKind, mockRewriter) in requestedRewriters)
            {
                mockRewriter.Verify(m => m.Rewrite(), Times.Never);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsFalseIfNotParsingAllowed()
        {
            var rewriteSession = RewriteSession(session => false, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            var actual = rewriteSession.TryRewrite();
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRewriteOnAnCheckedOutRewriterIfNotParsingAllowed()
        {
            var rewriteSession = RewriteSession(session => false, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var otherModule = new QualifiedModuleName("TestProject", string.Empty, "OtherTestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(otherModule);

            var requestedRewriters = mockRewriterProvider.RequestedRewriters();

            rewriteSession.TryRewrite();

            foreach (var (qmn, codeKind, mockRewriter) in requestedRewriters)
            {
                mockRewriter.Verify(m => m.Rewrite(), Times.Never);
            }
        }

        [Test]
        [Category("Rewriter")]
        public void ReturnsTheSameRewriterOnMultipleCheckoutsForTheSameModule()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var initialRewriter = rewriteSession.CheckOutModuleRewriter(module);
            var nextRewriter = rewriteSession.CheckOutModuleRewriter(module);

            Assert.AreEqual(initialRewriter, nextRewriter);
        }

        [Test]
        [Category("Rewriter")]
        public void CallsRewriteOnlyOnceForRewritersCheckedOutMultipleTimes()
        {
            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(module);
            var (qmn, codeKind, mockRewriter) = mockRewriterProvider.RequestedRewriters().Single();
            
            rewriteSession.TryRewrite();

            mockRewriter.Verify(m => m.Rewrite(), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void ChecksOutRewritersForTheTargetCodeKind()
        {
            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);

            var expectedCodeKind = rewriteSession.TargetCodeKind;
            var (qmn, actualCodeKind, mockRewriter) = mockRewriterProvider.RequestedRewriters().Single();
            Assert.AreEqual(expectedCodeKind, actualCodeKind);
        }

        protected IRewriteSession RewriteSession(Func<IRewriteSession, bool> rewritingAllowed,
            out MockRewriterProvider mockProvider)
        {
            var parseManager = new Mock<IParseManager>();
            parseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Callback((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => suspendAction())
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => SuspensionResult.Completed);
            return RewriteSession(parseManager.Object, rewritingAllowed, out mockProvider);
        }

        protected abstract IRewriteSession RewriteSession(IParseManager parseManager, Func<IRewriteSession, bool> rewritingAllowed, out MockRewriterProvider mockProvider);
    }


    public class MockRewriterProvider: IRewriterProvider
    {
        private readonly List<(QualifiedModuleName module, CodeKind codeKind, Mock<IExecutableModuleRewriter> moduleRewriter)> _requestedRewriters = new List<(QualifiedModuleName module, CodeKind codeKind, Mock<IExecutableModuleRewriter> moduleRewriter)>();

        public IExecutableModuleRewriter CodePaneModuleRewriter(QualifiedModuleName module)
        {
            var rewriter = CreateMockModuleRewriter();
            _requestedRewriters.Add((module, CodeKind.CodePaneCode, rewriter));
            return rewriter.Object;
        }

        private static Mock<IExecutableModuleRewriter> CreateMockModuleRewriter()
        {
            var mock = new Mock<IExecutableModuleRewriter>();
            mock.Setup(m => m.Rewrite());

            return mock;
        }

        public IExecutableModuleRewriter AttributesModuleRewriter(QualifiedModuleName module)
        {
            var rewriter = CreateMockModuleRewriter();
            _requestedRewriters.Add((module, CodeKind.AttributesCode, rewriter));
            return rewriter.Object;
        }

        public IEnumerable<(QualifiedModuleName module, CodeKind codeKind, Mock<IExecutableModuleRewriter> moduleRewriter)> RequestedRewriters()
        {
            return _requestedRewriters;
        }
    }
}