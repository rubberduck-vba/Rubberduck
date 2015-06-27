using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class UnsyncedCommitsPresenterTests
    {
        private Mock<ISourceControlProvider> _provider;
        private Mock<IUnsyncedCommitsView> _view;

        private IBranch _initialBranch;
        private UnsyncedCommitsPresenter _presenter;

        [TestInitialize]
        public void Intialize()
        {
            _initialBranch = new Branch("master", "refs/Heads/master", false, true);

            var incoming = new List<ICommit>() { new Commit("b9001d22", "Christopher J. McClellan", "Fixed the bazzle.") };
            var outgoing = new List<ICommit>() { new Commit("6129ebf7", "Mathieu Guindon", "Grammar fix.") };

            _provider = new Mock<ISourceControlProvider>();
            _provider.SetupGet(git => git.CurrentBranch).Returns(_initialBranch);
            _provider.SetupGet(git => git.UnsyncedRemoteCommits).Returns(incoming);
            _provider.SetupGet(git => git.UnsyncedLocalCommits).Returns(outgoing);

            _view = new Mock<IUnsyncedCommitsView>();
            _view.SetupProperty(v => v.CurrentBranch, String.Empty);
            _view.SetupProperty(v => v.IncomingCommits);
            _view.SetupProperty(v => v.OutgoingCommits);

            _presenter = new UnsyncedCommitsPresenter(_view.Object) { Provider = _provider.Object };
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_ViewBranchIsCurrentBranch()
        {
            //Arrange

            //Act
            _presenter.Refresh();

            //Assert
            Assert.AreEqual(_initialBranch.Name, _view.Object.CurrentBranch);
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_IncomingCommitsAreDisplayed()
        {
            //Arrange

            //Act
            _presenter.Refresh();

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), _view.Object.IncomingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_OutgoingCommitsAreDisplayed()
        {
            //Arrange

            //Act
            _presenter.Refresh();

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedLocalCommits.ToList(), _view.Object.OutgoingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnFetch_ProviderFetches()
        {
            //Arrange

            //Act - Simulate Fetch click
            _view.Raise(v => v.Fetch += null, EventArgs.Empty);

            //Assert
            _provider.Verify(git => git.Fetch(It.IsAny<string>()));
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterFetch_IncomingCommitsRefreshes()
        {
            //Arrange
            _provider.SetupGet(git => git.UnsyncedRemoteCommits)
                .Returns(new List<ICommit>() { new Commit("1111111", "Hosch250", "Fixed the foobarred bazzer.") });

            //Act - Simulate Fetch click
            _view.Raise(v => v.Fetch += null, EventArgs.Empty);

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), _view.Object.IncomingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnPull_ProviderPulls()
        {
            //Arrange
            //Act
            _view.Raise(v => v.Pull += null, EventArgs.Empty);
            //Assert
            _provider.Verify(git => git.Pull());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnPush_ProviderPushes()
        {
            //Arrange
            //Act
            _view.Raise(v => v.Push += null, EventArgs.Empty);
            //Assert
            _provider.Verify(git => git.Push());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnSync_ProviderPullsThenPushes()
        {
            //Arrange
            //Act
            _view.Raise(v => v.Sync += null, EventArgs.Empty);
            //Assert
            _provider.Verify(git => git.Pull());
            _provider.Verify(git => git.Push());
        }
    }
}
