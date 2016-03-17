using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class ChangesViewModelTests
    {
        private Mock<ISourceControlProvider> _providerMock;

        [TestInitialize]
        public void SetupMocks()
        {
            _providerMock = new Mock<ISourceControlProvider>();
            var branch = new Branch("master", "refs/Heads/master", false, true, null);
            _providerMock.SetupGet(git => git.CurrentBranch).Returns(branch);
        }

        [TestMethod]
        public void ProviderCommitIsCalledOnCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderStagesBeforeCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _providerMock.Verify(git => git.Stage(It.IsAny<IEnumerable<string>>()));
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderCommitsAndPushes()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitAction = CommitAction.CommitAndPush,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
            _providerMock.Verify(git => git.Push());
        }

        [TestMethod]
        public void ProviderCommitsAndSyncs()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitAction = CommitAction.CommitAndSync,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
            _providerMock.Verify(git => git.Pull());
            _providerMock.Verify(git => git.Push());
        }

        [TestMethod]
        public void RefreshDisplaysChangedFiles()
        {
            //arrange
            var fileStatusEntries = new List<FileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified),
                        new FileStatusEntry(@"C:\path\to\class.cls", FileStatus.Unaltered),
                        new FileStatusEntry(@"C:\path\to\added.bas", FileStatus.Added | FileStatus.Modified),
                        new FileStatusEntry(@"C:\path\to\addedUnmodified.bas", FileStatus.Added),
                        new FileStatusEntry(@"C:\path\to\untracked.frx", FileStatus.Untracked)
                    };

            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitAction = CommitAction.CommitAndSync
            };
            _providerMock.Setup(git => git.Status()).Returns(fileStatusEntries);

            //act
            vm.RefreshView();

            //Assert
            Assert.AreEqual(3, vm.IncludedChanges.Count, "Incorrect Included Changes");
            Assert.AreEqual(@"C:\path\to\untracked.frx", vm.UntrackedFiles[0].FilePath);
        }

        [TestMethod]
        public void CommitEnabledAfterActionSelectedAndMessageEntered()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //assert
            Assert.IsTrue(vm.CommitCommand.CanExecute(null));
        }

        [TestMethod]
        public void ClearCommitMessageAfterSuccessfulCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            Assert.AreEqual(string.Empty, vm.CommitMessage);
        }

        [TestMethod]
        public void RefreshChangesAfterCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            Assert.IsTrue(vm.IncludedChanges.Any());

            //act
            vm.CommitCommand.Execute(null);
            _providerMock.Setup(git => git.Status()).Returns(new List<FileStatusEntry>());

            //assert
            Assert.IsFalse(vm.IncludedChanges.Any());
        }

        [TestMethod]
        public void ExcludedIsNotClearedAfterRefresh()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                ExcludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            Assert.IsTrue(vm.ExcludedChanges.Any());

            //act
            vm.RefreshView();

            //assert
            Assert.IsTrue(vm.ExcludedChanges.Any());
        }

        [TestMethod]
        public void ExcludeFileExcludesFile()
        {
            //arrange
            var fileStatusEntries = new List<FileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified),
                        new FileStatusEntry(@"C:\path\to\class.cls", FileStatus.Unaltered),
                        new FileStatusEntry(@"C:\path\to\added.bas", FileStatus.Added | FileStatus.Modified),
                        new FileStatusEntry(@"C:\path\to\addedUnmodified.bas", FileStatus.Added),
                        new FileStatusEntry(@"C:\path\to\untracked.frx", FileStatus.Untracked)
                    };

            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitAction = CommitAction.CommitAndSync
            };
            _providerMock.Setup(git => git.Status()).Returns(fileStatusEntries);

            //act
            vm.ExcludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());

            //Assert
            Assert.AreEqual(1, vm.ExcludedChanges.Count);
        }

        // I need to figure out how to make this throw.
        [Ignore]
        [TestMethod]
        public void ChangesPresenter_WhenCommitFails_ActionFailedEventIsRaised()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _providerMock.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Untracked)
                    }
            };

            var wasRaised = false;

            vm.ErrorThrown += (e, sender) => wasRaised = true;

            //act
            vm.CommitCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
