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
        public void IsValidAtStart()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            Assert.AreEqual(RewriteSessionState.Valid, rewriteSession.Status);
        }

        [Test]
        [Category("Rewriter")]
        public void IsNotValidAfterSettingTheStatusToAnInvalidState()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            rewriteSession.Status = RewriteSessionState.StaleParseTree;
            Assert.AreNotEqual(RewriteSessionState.Valid,rewriteSession.Status);
        }

        [Test]
        [Category("Rewriter")]
        public void StaysNotValidAfterSettingTheStatusToAnInvalidState()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
            rewriteSession.Status = RewriteSessionState.Valid;
            Assert.AreNotEqual(RewriteSessionState.Valid, rewriteSession.Status);
        }

        [Test]
        [Category("Rewriter")]
        public void TheInvalidationStatusCannotBeChanged()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            rewriteSession.Status = RewriteSessionState.RewriteApplied;
            rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
            Assert.AreEqual(RewriteSessionState.RewriteApplied, rewriteSession.Status);
        }

        [Test]
        [Category("Rewriter")]
        public void StatusChangesToInvalidStateStaleParseTreeIfADirtyRewriterGetsCheckedOut()
        {
            var rewriteSession = RewriteSession(session =>
            {
                return true;
            }, out _, rewritersAreDirty: true);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            rewriteSession.CheckOutModuleRewriter(module);
            Assert.AreEqual(RewriteSessionState.StaleParseTree, rewriteSession.Status);
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
        public void TryRewriteReturnsTrueIfNoRewritersHaveBeenCheckedOut()
        {
            var rewriteSession = RewriteSession(session => true, out _);
            var actual = rewriteSession.TryRewrite();
            Assert.IsTrue(actual);
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
            rewriteSession.Status = RewriteSessionState.StaleParseTree;
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
            rewriteSession.Status = RewriteSessionState.StaleParseTree;
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
            rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
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

        [Test]
        [Category("Rewriter")]
        public void CallsRecoverSelectionOnNextParseOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.RecoverSavedSelectionsOnNextParse());

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.RecoverSavedSelectionsOnNextParse(), Times.Once);
        }

        [Test]
        [Category("Rewriter")]
        public void SavesSelectionForCheckedOutModulesOnRewrite()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.SaveSelections(It.IsAny<IEnumerable<QualifiedModuleName>>()));

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var module2 = new QualifiedModuleName("TestProject", string.Empty, "TestModule2");

            rewriteSession.CheckOutModuleRewriter(module);
            rewriteSession.CheckOutModuleRewriter(module2);
            rewriteSession.TryRewrite();
            
            selectionRecovererMock.Verify(m => m.SaveSelections(It.Is<IEnumerable<QualifiedModuleName>>(modules => modules.Count() == 2 && modules.Contains(module) && modules.Contains(module2))));
        }

        [Test]
        [Category("Rewriter")]
        public void AdjustsSelectionForCheckedOutModulesOnRewriteWhoseRewriterHasASelectionOffsetSet()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.AdjustSavedSelection(It.IsAny<QualifiedModuleName>(), It.IsAny<Selection>()));

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var module2 = new QualifiedModuleName("TestProject", string.Empty, "TestModule2");

            var selectionOffset = new Selection(2,3);

            var rewriter = rewriteSession.CheckOutModuleRewriter(module);
            rewriter.SelectionOffset = selectionOffset;
            rewriteSession.CheckOutModuleRewriter(module2);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.AdjustSavedSelection(module, selectionOffset), Times.Once);
            selectionRecovererMock.Verify(m => m.AdjustSavedSelection(module2, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriter")]
        public void ReplacesSelectionForCheckedOutModulesOnRewriteWhoseRewriterHasASelectionSet()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            var moduleSelection = new Selection(4,3);
            selectionRecovererMock.Setup(m => m.ReplaceSavedSelection(It.IsAny<QualifiedModuleName>(), It.IsAny<Selection>())).Callback((QualifiedModuleName qmn, Selection selection) => moduleSelection = selection);
            selectionRecovererMock.Setup(m => m.AdjustSavedSelection(It.IsAny<QualifiedModuleName>(), It.IsAny<Selection>())).Callback((QualifiedModuleName qmn, Selection selection) => moduleSelection = moduleSelection.Offset(selection));

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");

            var selectionReplacement = new Selection(2, 3);
            var selectionOffset = new Selection(42,42);

            var rewriter = rewriteSession.CheckOutModuleRewriter(module);
            rewriter.Selection = selectionReplacement;
            rewriter.SelectionOffset = selectionOffset;
            rewriteSession.TryRewrite();

            Assert.AreEqual(selectionReplacement, moduleSelection);
        }

        [Test]
        [Category("Rewriter")]
        public void FirstAppliesOffsetAndThenReplacesSelectionForCheckedOutModulesOnRewriteWhoseRewriterHasASelectionAndASelectionOffsetSet()
        {
            var selectionRecovererMock = new Mock<ISelectionRecoverer>();
            selectionRecovererMock.Setup(m => m.ReplaceSavedSelection(It.IsAny<QualifiedModuleName>(), It.IsAny<Selection>()));

            var rewriteSession = RewriteSession(session => true, out var mockRewriterProvider, selectionRecoverer: selectionRecovererMock.Object);
            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var module2 = new QualifiedModuleName("TestProject", string.Empty, "TestModule2");

            var selectionReplacement = new Selection(2, 3);

            var rewriter = rewriteSession.CheckOutModuleRewriter(module);
            rewriter.Selection = selectionReplacement;
            rewriteSession.CheckOutModuleRewriter(module2);
            rewriteSession.TryRewrite();

            selectionRecovererMock.Verify(m => m.ReplaceSavedSelection(module, selectionReplacement), Times.Once);
            selectionRecovererMock.Verify(m => m.ReplaceSavedSelection(module2, It.IsAny<Selection>()), Times.Never);
        }


        protected IExecutableRewriteSession RewriteSession(Func<IRewriteSession, bool> rewritingAllowed,
            out MockRewriterProvider mockProvider, bool rewritersAreDirty = false, ISelectionRecoverer selectionRecoverer = null)
        {
            var parseManager = new Mock<IParseManager>();
            parseManager.Setup(m => m.OnSuspendParser(It.IsAny<object>(), It.IsAny<IEnumerable<ParserState>>(), It.IsAny<Action>(), It.IsAny<int>()))
                .Callback((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => suspendAction())
                .Returns((object requestor, IEnumerable<ParserState> allowedStates, Action suspendAction, int timeout) => new SuspensionResult(SuspensionOutcome.Completed));
            return RewriteSession(parseManager.Object, rewritingAllowed, out mockProvider, rewritersAreDirty, selectionRecoverer);
        }

        protected abstract IExecutableRewriteSession RewriteSession(IParseManager parseManager, Func<IRewriteSession, bool> rewritingAllowed, out MockRewriterProvider mockProvider, bool rewritersAreDirty = false, ISelectionRecoverer selectionRecoverer = null);
    }


    public class MockRewriterProvider: IRewriterProvider
    {
        private readonly List<(QualifiedModuleName module, CodeKind codeKind, Mock<IExecutableModuleRewriter> moduleRewriter)> _requestedRewriters = new List<(QualifiedModuleName module, CodeKind codeKind, Mock<IExecutableModuleRewriter> moduleRewriter)>();

        private readonly bool _createdRewritersAreDirty;

        public MockRewriterProvider(bool createdRewritersAreDirty = false)
        {
            _createdRewritersAreDirty = createdRewritersAreDirty;
        }

        public IExecutableModuleRewriter CodePaneModuleRewriter(QualifiedModuleName module)
        {
            var rewriter = CreateMockModuleRewriter();
            _requestedRewriters.Add((module, CodeKind.CodePaneCode, rewriter));
            return rewriter.Object;
        }

        private Mock<IExecutableModuleRewriter> CreateMockModuleRewriter()
        {
            var mock = new Mock<IExecutableModuleRewriter>();
            mock.Setup(m => m.Rewrite());
            mock.Setup(m => m.IsDirty).Returns(_createdRewritersAreDirty);
            mock.SetupProperty(m => m.SelectionOffset);
            mock.SetupProperty(m => m.Selection);

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