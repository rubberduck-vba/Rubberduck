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
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => new SuspensionResult(SuspensionOutcome.Completed));

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
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => new SuspensionResult(SuspensionOutcome.Completed));

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
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => new SuspensionResult(SuspensionOutcome.UnexpectedError));

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
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => new SuspensionResult(SuspensionOutcome.Completed));

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

        [Test]
        [Category("Rewriter")]
        public void CallsRecoverOpenStateOnNextParseOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.RecoverOpenStateOnNextParse());

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.RecoverOpenStateOnNextParse(), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void SavesOpenStateForCheckedOutModulesOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.SaveOpenState(It.IsAny<IEnumerable<QualifiedModuleName>>()));

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var module2 = new QualifiedModuleName("TestProject", string.Empty, "TestModule2");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(module2);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.SaveOpenState(It.Is<IEnumerable<QualifiedModuleName>>(modules => modules.Count() == 2 && modules.Contains(module) && modules.Contains(module2))));
        }

        [Test]
        [Category("Rewriter")]
        public void SavesOpenStateBeforeRecoverOpenStateOnNextParseOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            var lastOperation = string.Empty;
            selectionRecovererMock.Setup(m => m.SaveOpenState(It.IsAny<IEnumerable<QualifiedModuleName>>())).Callback(() => lastOperation = "SaveOpenState");
            selectionRecovererMock.Setup(m => m.RecoverOpenStateOnNextParse()).Callback(() => lastOperation = "RecoverOpenStateOnNextParse");

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            Assert.AreEqual("RecoverOpenStateOnNextParse", lastOperation);
        }

        [Test]
        [Category("Rewriter")]
        public void CallsRecoverActiveCodePaneOnNextParseOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.RecoverActiveCodePaneOnNextParse());

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.RecoverActiveCodePaneOnNextParse(), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void SavesActiveCodePaneOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.SaveActiveCodePane());

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.SaveActiveCodePane(), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void SavesActiveCodePaneBeforeRestoringItOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            var lastOperation = string.Empty;
            selectionRecovererMock.Setup(m => m.SaveActiveCodePane()).Callback(() => lastOperation = "SaveActiveCodePane");
            selectionRecovererMock.Setup(m => m.RecoverActiveCodePaneOnNextParse()).Callback(() => lastOperation = "RecoverActiveCodePaneOnNextParse");

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            Assert.AreEqual("RecoverActiveCodePaneOnNextParse", lastOperation);
        }

        protected override IExecutableRewriteSession RewriteSession(IParseManager parseManager, Func<IRewriteSession, bool> rewritingAllowed, out MockRewriterProvider mockProvider, bool rewritersAreDirty = false, ISelectionRecoverer selectionRecoverer = null)
        {
            if (selectionRecoverer == null)
            {
                selectionRecoverer = new Mock<ISelectionRecoverer>().Object;
            }
            mockProvider = new MockRewriterProvider(rewritersAreDirty);
            return new AttributesRewriteSession(parseManager, mockProvider, selectionRecoverer, rewritingAllowed);
        }
    }
}