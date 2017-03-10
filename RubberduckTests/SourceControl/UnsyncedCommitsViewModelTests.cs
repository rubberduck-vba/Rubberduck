using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class UnsyncedCommitsViewModelTests
    {
        private Mock<ISourceControlProvider> _provider;

        private IBranch _initialBranch;

        [TestInitialize]
        public void Intialize()
        {
            var masterRemote = new Mock<LibGit2Sharp.Branch>();
            masterRemote.SetupGet(git => git.Tip).Returns(new Mock<LibGit2Sharp.Commit>().Object);
            masterRemote.SetupGet(git => git.FriendlyName).Returns("master");

            _initialBranch = new Branch("master", "refs/Heads/master", false, true, masterRemote.Object);

            var incoming = new List<ICommit> { new Commit("b9001d22", "Christopher J. McClellan", "Fixed the bazzle.") };
            var outgoing = new List<ICommit> { new Commit("6129ebf7", "Mathieu Guindon", "Grammar fix.") };

            _provider = new Mock<ISourceControlProvider>();
            _provider.SetupGet(git => git.CurrentBranch).Returns(_initialBranch);
            _provider.SetupGet(git => git.UnsyncedRemoteCommits).Returns(incoming);
            _provider.SetupGet(git => git.UnsyncedLocalCommits).Returns(outgoing);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_ViewBranchIsCurrentBranch()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            Assert.AreEqual(_initialBranch.Name, vm.CurrentBranch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_IncomingCommitsAreDisplayed()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), vm.IncomingCommits.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_OutgoingCommitsAreDisplayed()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedLocalCommits.ToList(), vm.OutgoingCommits.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_OnFetch_ProviderFetches()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            vm.FetchCommitsCommand.Execute(null);

            _provider.Verify(git => git.Fetch(It.IsAny<string>()));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_AfterFetch_IncomingCommitsRefreshes()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.SetupGet(git => git.UnsyncedRemoteCommits)
                .Returns(new List<ICommit> { new Commit("1111111", "Hosch250", "Fixed the foobarred bazzer.") });

            vm.FetchCommitsCommand.Execute(null);

            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), vm.IncomingCommits.ToList());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_OnPull_ProviderPulls()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            vm.PullCommitsCommand.Execute(null);

            _provider.Verify(git => git.Pull());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_OnPush_ProviderPushes()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            vm.PushCommitsCommand.Execute(null);

            _provider.Verify(git => git.Push());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_OnSync_ProviderPullsThenPushes()
        {
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            vm.SyncCommitsCommand.Execute(null);

            _provider.Verify(git => git.Pull());
            _provider.Verify(git => git.Push());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_WhenFetchFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.Setup(p => p.Fetch(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.FetchCommitsCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_WhenPushFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.Setup(p => p.Push())
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.PushCommitsCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_WhenPullFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.Setup(p => p.Pull())
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.PullCommitsCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UnsyncedPresenter_WhenSyncFails_ActionFailedEventIsRaised()
        {
            var wasRaised = false;
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.Setup(p => p.Pull())
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            vm.SyncCommitsCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
