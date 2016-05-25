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

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_ViewBranchIsCurrentBranch()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Assert
            Assert.AreEqual(_initialBranch.Name, vm.CurrentBranch);
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_IncomingCommitsAreDisplayed()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), vm.IncomingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterRefresh_OutgoingCommitsAreDisplayed()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedLocalCommits.ToList(), vm.OutgoingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnFetch_ProviderFetches()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Act - Simulate Fetch click
            vm.FetchCommitsCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Fetch(It.IsAny<string>()));
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterFetch_IncomingCommitsRefreshes()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            _provider.SetupGet(git => git.UnsyncedRemoteCommits)
                .Returns(new List<ICommit> { new Commit("1111111", "Hosch250", "Fixed the foobarred bazzer.") });

            //Act - Simulate Fetch click
            vm.FetchCommitsCommand.Execute(null);

            //Assert
            CollectionAssert.AreEquivalent(_provider.Object.UnsyncedRemoteCommits.ToList(), vm.IncomingCommits.ToList());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnPull_ProviderPulls()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Act
            vm.PullCommitsCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Pull());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnPush_ProviderPushes()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Act
            vm.PushCommitsCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Push());
        }

        [TestMethod]
        public void UnsyncedPresenter_OnSync_ProviderPullsThenPushes()
        {
            //Arrange
            var vm = new UnsyncedCommitsViewViewModel
            {
                Provider = _provider.Object
            };

            //Act
            vm.SyncCommitsCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Pull());
            _provider.Verify(git => git.Push());
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenFetchFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.FetchCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenPushFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.PushCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenPullFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.PullCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenSyncFails_ActionFailedEventIsRaised()
        {
            //arrange
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

            //act
            vm.SyncCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
