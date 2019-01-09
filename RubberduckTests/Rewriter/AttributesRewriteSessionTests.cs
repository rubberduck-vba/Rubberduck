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
    public class AttributesRewriteSessionTests : RewriteSessionTestBase
    {
        [Test]
        [Category("Rewriter")]
        public void UsesASuspendActionToRewrite()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Callback((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => suspendAction())
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => SuspensionResult.Completed);

            var rewriteSession = RewriteSession(mockParseManager.Object, session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);

            rewriteSession.TryRewrite();

            mockParseManager.Verify(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void DoesNotCallRewriteOutsideTheSuspendAction()
        {
            var mockParseManager = new Mock<IParseManager>();
            mockParseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => SuspensionResult.Completed);

            var rewriteSession = RewriteSession(mockParseManager.Object, session => true, out var mockRewriterProvider);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            var (qmn, codeKind, mockRewriter) = mockRewriterProvider.RequestedRewriters().Single();

            rewriteSession.TryRewrite();

            mockRewriter.Verify(m => m.Rewrite(), Times.Never);
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsFalseIfNotInvalidatedAndParsingAllowedAndSuspensionDoesNotComplete()
        {
            var parseManager = new Mock<IParseManager>();
            parseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Callback((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => suspendAction())
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => SuspensionResult.UnexpectedError);

            var rewriteSession = RewriteSession(parseManager.Object, session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            var actual = rewriteSession.TryRewrite();
            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void TryRewriteReturnsTrueIfNotInvalidatedAndParsingAllowedAndSuspensionCompletes()
        {
            var parseManager = new Mock<IParseManager>();
            parseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Callback((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => suspendAction())
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => SuspensionResult.Completed);

            var rewriteSession = RewriteSession(parseManager.Object, session => true, out _);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            var actual = rewriteSession.TryRewrite();
            Assert.IsTrue(actual);
        }

        [Test]
        [Category("Rewriter")]
        public void TargetsAttributesCode()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            Assert.AreEqual(CodeKind.AttributesCode, rewriteSession.TargetCodeKind);
        }

        protected override IRewriteSession RewriteSession(IParseManager parseManager, Func<IRewriteSession, bool> rewritingAllowed, out MockRewriterProvider mockProvider, bool rewritersAreDirty = false)
        {
            mockProvider = new MockRewriterProvider(rewritersAreDirty);
            return new AttributesRewriteSession(parseManager, mockProvider, rewritingAllowed);
        }
    }
}