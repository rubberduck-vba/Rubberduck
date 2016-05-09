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
        private Mock<ISourceControlProvider> _provider;

        [TestInitialize]
        public void SetupMocks()
        {
            _provider = new Mock<ISourceControlProvider>();
            var branch = new Branch("master", "refs/Heads/master", false, true, null);
            _provider.SetupGet(git => git.CurrentBranch).Returns(branch);
        }

        [TestMethod]
        public void ProviderCommitIsCalledOnCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _provider.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderStagesBeforeCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            _provider.Verify(git => git.Stage(It.IsAny<IEnumerable<string>>()));
            _provider.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderCommits_NotificationOnSuccess()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            var errorThrown = bool.FalseString; // need a reference type

            vm.ErrorThrown += (sender, e) =>
            {
                lock (errorThrown)
                {
                    MultiAssert.Aggregate(
                        () => Assert.AreEqual(e.Message, Rubberduck.UI.RubberduckUI.SourceControl_CommitStatus),
                        () =>
                            Assert.AreEqual(e.InnerMessage,
                                Rubberduck.UI.RubberduckUI.SourceControl_CommitStatus_CommitSuccess),
                        () => Assert.AreEqual(e.NotificationType, NotificationType.Info));

                    errorThrown = bool.TrueString;
                }
            };

            //act
            vm.CommitCommand.Execute(null);

            //assert
            lock (errorThrown)
            {
                Assert.IsTrue(bool.Parse(errorThrown));
            }
        }

        [TestMethod]
        public void ProviderCommitsAndPushes()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
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
            _provider.Verify(git => git.Commit(It.IsAny<string>()));
            _provider.Verify(git => git.Push());
        }

        [TestMethod]
        public void ProviderCommitsAndSyncs()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
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
            _provider.Verify(git => git.Commit(It.IsAny<string>()));
            _provider.Verify(git => git.Pull());
            _provider.Verify(git => git.Push());
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
                Provider = _provider.Object,
                CommitAction = CommitAction.CommitAndSync
            };
            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);

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
                Provider = _provider.Object,
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
        public void RefreshChangesAfterCommit()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
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
            _provider.Setup(git => git.Status()).Returns(new List<FileStatusEntry>());

            //assert
            Assert.IsFalse(vm.IncludedChanges.Any());
        }

        [TestMethod]
        public void ExcludedIsNotClearedAfterRefresh()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
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
                Provider = _provider.Object,
                CommitAction = CommitAction.CommitAndSync
            };
            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);

            //act
            vm.ExcludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());

            //Assert
            Assert.AreEqual(1, vm.ExcludedChanges.Count);
        }

        [TestMethod]
        public void ChangesPresenter_WhenCommitFails_ActionFailedEventIsRaised()
        {
            //arrange
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                CommitMessage = "Test Message",
                CommitAction = CommitAction.Commit,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Untracked)
                    }
            };

            _provider.Setup(p => p.Commit(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            var wasRaised = false;

            vm.ErrorThrown += (e, sender) => wasRaised = true;

            //act
            vm.CommitCommand.Execute(null);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void Undo_UndoesChanges()
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
                Provider = _provider.Object
            };

            var localLocation = "C:\\users\\desktop\\git\\";

            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);
            _provider.SetupGet(git => git.CurrentRepository).Returns(new Repository{LocalLocation = localLocation});

            //act
            vm.UndoChangesToolbarButtonCommand.Execute(fileStatusEntries[0]);

            //Assert
            _provider.Verify(git => git.Undo(localLocation + fileStatusEntries[0].FilePath));
        }

        [TestMethod]
        public void IncludeChanges_AddsUntrackedFile()
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

            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);

            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object
            };

            //act
            vm.IncludeChangesToolbarButtonCommand.Execute(fileStatusEntries.Last());

            //Assert
            _provider.Verify(git => git.AddFile(fileStatusEntries.Last().FilePath));
        }

        [TestMethod]
        public void IncludeChanges_IncludesExcludedFile()
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

            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                CommitAction = CommitAction.CommitAndSync
            };

            //act-assert
            vm.ExcludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());
            Assert.AreEqual(1, vm.ExcludedChanges.Count);

            //act-assert
            vm.IncludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());
            Assert.AreEqual(3, vm.IncludedChanges.Count);
            Assert.AreEqual(0, vm.ExcludedChanges.Count);
        }

        [TestMethod]
        public void UndoFails_ActionFailedEventIsRaised()
        {
            //arrange
            var wasRaised = false;
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
                Provider = _provider.Object
            };

            var localLocation = "C:\\users\\desktop\\git\\";

            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);
            _provider.SetupGet(git => git.CurrentRepository).Returns(new Repository { LocalLocation = localLocation });

            _provider.Setup(p => p.Undo(It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            vm.ErrorThrown += (sender, error) => wasRaised = true;

            //act
            vm.UndoChangesToolbarButtonCommand.Execute(fileStatusEntries[0]);

            //assert
            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
