using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SourceControl;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class ChangesPresenterTests
    {
        private Mock<IChangesView> _viewMock;
        private Mock<ISourceControlProvider> _providerMock;

        [TestInitialize]
        public void SetupMocks()
        {
            _viewMock = new Mock<IChangesView>();
            _providerMock = new Mock<ISourceControlProvider>();
        }

        [TestMethod]
        public void ProviderCommitIsCalledOnCommit()
        {
            //arrange
            _viewMock.SetupProperty(v => v.IncludedChanges);
            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);

            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\file.txt", FileStatus.Modified) };

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderStagesBeforeCommit()
        {
            //arrange
            _viewMock.SetupProperty(v => v.IncludedChanges);
            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);

            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\file.txt", FileStatus.Modified) };

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());

            //assert
            _providerMock.Verify(git => git.Stage(It.IsAny<IEnumerable<string>>()));
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
        }

        [TestMethod]
        public void ProviderCommitsAndPushes()
        {
            //arrange
            _viewMock.SetupProperty(v => v.IncludedChanges);
            _viewMock.SetupProperty(v => v.CommitAction, CommitAction.CommitAndPush);

            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);
            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\file.txt", FileStatus.Modified) };

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
            _providerMock.Verify(git => git.Push());
        }

        [TestMethod]
        public void ProviderCommitsAndSyncs()
        {
            //arrange
            _viewMock.SetupProperty(v => v.IncludedChanges);
            _viewMock.SetupProperty(v => v.CommitAction, CommitAction.CommitAndSync);

            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);
            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\file.txt", FileStatus.Modified) };

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());

            //assert
            _providerMock.Verify(git => git.Commit(It.IsAny<string>()));
            _providerMock.Verify(git => git.Pull());
            _providerMock.Verify(git => git.Push());
        }

        [TestMethod]
        public void RefreshDisplaysChangedFiles()
        {
            //arrange
            var fileStatusEntries = new List<FileStatusEntry>()
            {
                new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified),
                new FileStatusEntry(@"C:\path\to\class.cls", FileStatus.Unaltered),
                new FileStatusEntry(@"C:\path\to\added.bas", FileStatus.Added | FileStatus.Modified),
                new FileStatusEntry(@"C:\path\to\untracked.frx", FileStatus.Untracked)
            };

            _viewMock.SetupProperty(v => v.IncludedChanges);
            _viewMock.SetupProperty(v => v.UntrackedFiles);
            _providerMock.Setup(git => git.Status()).Returns(fileStatusEntries);

            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);
            //act
            presenter.Refresh();

            //Assert
            Assert.AreEqual(2, _viewMock.Object.IncludedChanges.Count, "Incorrect Included Changes");
            Assert.AreEqual(@"C:\path\to\untracked.frx", _viewMock.Object.UntrackedFiles[0].FilePath);
        }

        [TestMethod]
        public void  CommitEnabledAfterActionSelectedAndMessageEntered()
        {
            //arrange
            _viewMock.SetupAllProperties();
            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);

            //act
            _viewMock.Object.CommitMessage = "Test Message";
            _viewMock.Raise(v => v.CommitMessageChanged += null, new EventArgs());

            _viewMock.Object.CommitAction = CommitAction.Commit;
            _viewMock.Raise(v => v.SelectedActionChanged += null, new EventArgs());

            //assert
            Assert.IsTrue(_viewMock.Object.CommitEnabled);
        }

        [TestMethod]
        public void ClearCommitMessageAfterSuccessfulCommit()
        {

            _viewMock.SetupAllProperties();
            _viewMock.Object.CommitMessage = "Test Commit";
            _viewMock.Object.CommitAction = CommitAction.Commit;
            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified) };

            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());

            //assert
            Assert.AreEqual(string.Empty, _viewMock.Object.CommitMessage);
        }

        [TestMethod]
        public void RefreshChangesAfterCommit()
        {
            //arrange
            _viewMock.SetupAllProperties();
            _viewMock.Object.CommitMessage = "Test Commit";
            _viewMock.Object.CommitAction = CommitAction.Commit;
            _viewMock.Object.IncludedChanges = new List<IFileStatusEntry>() {new FileStatusEntry(@"C:\path\to\module.bas",FileStatus.Modified) };

            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);

            //act
            _viewMock.Raise(v => v.Commit += null, new EventArgs());
            _providerMock.Setup(git => git.Status()).Returns(new List<FileStatusEntry>());
            
            //assert
            Assert.IsFalse(_viewMock.Object.IncludedChanges.Any());
        }

        [TestMethod]
        public void ExcludedIsClearedAfterRefresh()
        {
            //arrange
            _viewMock.SetupProperty(v => v.ExcludedChanges);
            var presenter = new ChangesPresenter(_viewMock.Object, _providerMock.Object);
            _viewMock.Object.ExcludedChanges = new List<IFileStatusEntry>() { new FileStatusEntry(@"C:\path\to\module.bas", FileStatus.Modified) };

            Assert.IsTrue(_viewMock.Object.ExcludedChanges.Any(), "No changes found prior to refresh. Issue with Test code.");

            //act
            presenter.Refresh();

            //
            Assert.IsFalse(_viewMock.Object.ExcludedChanges.Any());
        }
    }
}
