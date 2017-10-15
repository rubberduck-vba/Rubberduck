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
        private readonly object _locker = new object();

        [TestInitialize]
        public void SetupMocks()
        {
            _provider = new Mock<ISourceControlProvider>();
            var branch = new Branch("master", "refs/Heads/master", false, true, null);
            _provider.SetupGet(git => git.CurrentBranch).Returns(branch);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCommitIsCalledOnCommit()
        {
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            vm.CommitCommand.Execute(null);

            _provider.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderStagesBeforeCommit()
        {
            var vm = new ChangesViewViewModel
            {
                Provider = _provider.Object,
                IncludedChanges =
                    new ObservableCollection<IFileStatusEntry>
                    {
                        new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified)
                    }
            };

            vm.CommitCommand.Execute(null);

            _provider.Verify(git => git.Stage(It.IsAny<IEnumerable<string>>()));
            _provider.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCommits_NotificationOnSuccess()
        {
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
                lock (_locker)
                {
                    MultiAssert.Aggregate(
                        () => Assert.AreEqual(e.Title, Rubberduck.UI.RubberduckUI.SourceControl_CommitStatus),
                        () =>
                            Assert.AreEqual(e.InnerMessage,
                                Rubberduck.UI.RubberduckUI.SourceControl_CommitStatus_CommitSuccess),
                        () => Assert.AreEqual(e.NotificationType, NotificationType.Info));

                    errorThrown = bool.TrueString;
                }
            };

            vm.CommitCommand.Execute(null);

            lock (_locker)
            {
                Assert.IsTrue(bool.Parse(errorThrown));
            }
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCommitsAndPushes()
        {
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

            vm.CommitCommand.Execute(null);

            _provider.Verify(git => git.Commit(It.IsAny<string>()));
            _provider.Verify(git => git.Push());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ProviderCommitsAndSyncs()
        {
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

            vm.CommitCommand.Execute(null);

            _provider.Verify(git => git.Commit(It.IsAny<string>()));
            _provider.Verify(git => git.Pull());
            _provider.Verify(git => git.Push());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void RefreshDisplaysChangedFiles()
        {
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

            vm.RefreshView();

            Assert.AreEqual(3, vm.IncludedChanges.Count, "Incorrect Included Changes");
            Assert.AreEqual(@"C:\path\to\untracked.frx", vm.UntrackedFiles[0].FilePath);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void CommitEnabledAfterActionSelectedAndMessageEntered()
        {
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

            Assert.IsTrue(vm.CommitCommand.CanExecute(null));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void RefreshChangesAfterCommit()
        {
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

            vm.CommitCommand.Execute(null);
            _provider.Setup(git => git.Status()).Returns(new List<FileStatusEntry>());

            Assert.IsFalse(vm.IncludedChanges.Any());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ExcludedIsNotClearedAfterRefresh()
        {
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

            vm.RefreshView();

            Assert.IsTrue(vm.ExcludedChanges.Any());
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ExcludeFileExcludesFile()
        {
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

            vm.ExcludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());

            Assert.AreEqual(1, vm.ExcludedChanges.Count);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ChangesPresenter_WhenCommitFails_ActionFailedEventIsRaised()
        {
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

            vm.CommitCommand.Execute(null);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void Undo_UndoesChanges()
        {
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

            var localLocation = @"C:\users\desktop\git\";

            _provider.Setup(git => git.Status()).Returns(fileStatusEntries);
            _provider.SetupGet(git => git.CurrentRepository).Returns(new Repository{LocalLocation = localLocation});

            vm.UndoChangesToolbarButtonCommand.Execute(fileStatusEntries[0]);

            _provider.Verify(git => git.Undo(@"C:\users\desktop\git\module.bas"));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void IncludeChanges_AddsUntrackedFile()
        {
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

            vm.IncludeChangesToolbarButtonCommand.Execute(fileStatusEntries.Last());

            _provider.Verify(git => git.AddFile(fileStatusEntries.Last().FilePath));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void IncludeChanges_IncludesExcludedFile()
        {
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

            vm.ExcludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());
            Assert.AreEqual(1, vm.ExcludedChanges.Count);

            vm.IncludeChangesToolbarButtonCommand.Execute(fileStatusEntries.First());
            Assert.AreEqual(3, vm.IncludedChanges.Count);
            Assert.AreEqual(0, vm.ExcludedChanges.Count);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void UndoFails_ActionFailedEventIsRaised()
        {
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

            vm.UndoChangesToolbarButtonCommand.Execute(fileStatusEntries[0]);

            Assert.IsTrue(wasRaised, "ActionFailedEvent was not raised.");
        }
    }
}
