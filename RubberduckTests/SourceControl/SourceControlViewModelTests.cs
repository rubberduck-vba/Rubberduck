using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Security;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.SourceControl;
using Rubberduck.UI;
using Rubberduck.UI.SourceControl;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class SourceControlViewModelTests
    {
        private Mock<VBE> _vbe;
        private MockWindowsCollection _windows;
        private Mock<Window> _window;
#pragma warning disable 169
        private object _toolWindow;
#pragma warning restore 169

        private SourceControlViewViewModel _vm;

        private ChangesViewViewModel _changesVM;
        private BranchesViewViewModel _branchesVM;
        private UnsyncedCommitsViewViewModel _unsyncedVM;
        private SettingsViewViewModel _settingsVM;

        private Mock<ISourceControlConfigProvider> _configService;

        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;

        private Mock<ISourceControlProviderFactory> _providerFactory;
        private Mock<ISourceControlProvider> _provider;

        [TestInitialize]
        public void InitializeMocks()
        {
            _window = Mocks.MockFactory.CreateWindowMock();
            _windows = new MockWindowsCollection(new List<Window> { _window.Object });
            _vbe = Mocks.MockFactory.CreateVbeMock(_windows);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            _configService = new Mock<ISourceControlConfigProvider>();
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            _folderBrowser = new Mock<IFolderBrowser>();
            _folderBrowserFactory = new Mock<IFolderBrowserFactory>();
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>())).Returns(_folderBrowser.Object);
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>(), false)).Returns(_folderBrowser.Object);


            var masterRemote = new Mock<LibGit2Sharp.Branch>();
            masterRemote.SetupGet(git => git.Tip).Returns(new Mock<LibGit2Sharp.Commit>().Object);
            masterRemote.SetupGet(git => git.FriendlyName).Returns("master");

            var initialBranch = new Branch("master", "refs/Heads/master", false, true, masterRemote.Object);

            _provider = new Mock<ISourceControlProvider>();
            _provider.SetupGet(git => git.CurrentBranch).Returns(initialBranch);
            _provider.SetupGet(git => git.UnsyncedLocalCommits).Returns(new List<ICommit>());
            _provider.SetupGet(git => git.UnsyncedRemoteCommits).Returns(new List<ICommit>());
            _provider.Setup(git => git.InitVBAProject(It.IsAny<string>())).Returns(GetDummyRepo());
            _provider.Setup(git => git.Clone(It.IsAny<string>(), It.IsAny<string>())).Returns(GetDummyRepo());
            _provider.Setup(git => git.CurrentRepository).Returns(GetDummyRepo());

            _providerFactory = new Mock<ISourceControlProviderFactory>();
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>()))
                .Returns(_provider.Object);
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Returns(_provider.Object);
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<SecureCredentials>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Returns(_provider.Object);

            _changesVM = new ChangesViewViewModel();
            _branchesVM = new BranchesViewViewModel();
            _unsyncedVM = new UnsyncedCommitsViewViewModel();
            _settingsVM = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, new Rubberduck.UI.OpenFileDialog());
        }

        private void SetupValidVbProject()
        {
            var project = new Mock<VBProject>().SetupProperty(p => p.HelpFile, DummyRepoId);
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);
        }

        private void VerifyOffline()
        {
            Assert.AreEqual("Offline", _vm.Status);
        }

        private void VerifyChildPresentersHaveProviders()
        {
            Assert.IsNotNull(_settingsVM.Provider, "_settingsPresenter.Provider was null");
            Assert.IsNotNull(_branchesVM.Provider, "_branchesPresenter.Provider was null");
            Assert.IsNotNull(_changesVM.Provider, "_changesPresenter.Provider was null");
            Assert.IsNotNull(_unsyncedVM.Provider, "_unsyncedPresenter.Object.Provider was null");
        }

        private void SetupVM()
        {
            _vm = new SourceControlViewViewModel(_vbe.Object, new RubberduckParserState(), _providerFactory.Object, _folderBrowserFactory.Object,
                _configService.Object, new ChangesView(_changesVM), new BranchesView(_branchesVM),
                new UnsyncedCommitsView(_unsyncedVM), new SettingsView(_settingsVM), new CodePaneWrapperFactory(), new Mock<IMessageBox>().Object);
        }

        [TestMethod]
        public void StatusIsOfflineWhenNoRepoIsFoundInConfig()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(new SourceControlSettings());

            SetupValidVbProject();
            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineWhenRepoListIsEmpty()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(new SourceControlSettings());

            SetupValidVbProject();
            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineIfNoMatchingRepoExists()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(new SourceControlSettings());

            var project = new Mock<VBProject>().SetupProperty(p => p.Name, "FooBar");
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);

            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineWhenMultipleReposAreFound()
        {
            //arrange
            var config = GetDummyConfig();
            config.Repositories.Add(new Repository { Id = DummyRepoId });

            _configService.Setup(c => c.Create()).Returns(config);

            SetupValidVbProject();
            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            VerifyOffline();

        }

        [TestMethod]
        public void StatusIsOnlineWhenRepoIsFound()
        {
            //arrange 
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();
            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            Assert.AreEqual("Online", _vm.Status);
        }

        [TestMethod]
        public void ChildPresentersHaveValidProviderIfRepoIsFoundInConfig()
        {
            //arrange 
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();
            SetupVM();

            //act
            _vm.RefreshCommand.Execute(null);

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotAddedToConfig()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);
            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            _configService.Verify(c => c.Save(It.IsAny<SourceControlSettings>()), Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotCreated()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);
            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            _provider.Verify(git => git.InitVBAProject(It.IsAny<string>()), Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_RepoIsAddedToConfig()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            _configService.Verify(c => c.Save(It.IsAny<SourceControlSettings>()), Times.Once);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_RepoIsInitalized()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            _provider.Verify(git => git.InitVBAProject(It.IsAny<string>()), Times.Once);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserCancels_RepoIsNotAddedToConfig()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            SetupVM();

            //act
            _vm.OpenRepoCommand.Execute(null);

            //assert
            _configService.Verify(c => c.Save(It.IsAny<SourceControlSettings>()), Times.Never);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_RepoIsAddedToConfig()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.OpenRepoCommand.Execute(null);

            //assert
            _configService.Verify(c => c.Save(It.IsAny<SourceControlSettings>()), Times.Once);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_StatusIsOnline()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            Assert.AreEqual("Online", _vm.Status);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_StatusIsOnline()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.OpenRepoCommand.Execute(null);

            //assert
            Assert.AreEqual("Online", _vm.Status);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_ChildPresenterSourceControlProvidersAreSet()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.InitRepoCommand.Execute(null);

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_ChildPresenterSourceControlProvidersAreSet()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            SetupVM();

            //act
            _vm.OpenRepoCommand.Execute(null);

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void BranchesPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed";
            const string expectedMessage = "Details about failure.";

            _provider.Setup(p => p.Publish(It.IsAny<string>()))
                .Throws(
                    new SourceControlException(expectedTitle,
                        new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupVM();
            _vm.Provider = _provider.Object;

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _branchesVM.PublishBranchToolbarButtonCommand.Execute("");

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);

            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void ChangesPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed";
            const string expectedMessage = "Details about failure.";

            _provider.Setup(p => p.Commit(It.IsAny<string>()))
                .Throws(
                    new SourceControlException(expectedTitle,
                        new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupVM();
            _vm.Provider = _provider.Object;
            _changesVM.CommitMessage = "test";
            _changesVM.IncludedChanges = new ObservableCollection<IFileStatusEntry>()
            {
                new FileStatusEntry("path", FileStatus.Added)
            };

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _changesVM.CommitCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);
            
            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            _provider.Setup(p => p.Pull())
                .Throws(
                    new SourceControlException(expectedTitle,
                        new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupVM();
            _vm.Provider = _provider.Object;

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _unsyncedVM.PullCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);

            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenProviderCreationFails_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Throws(new SourceControlException(expectedTitle,
                    new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupValidVbProject();
            SetupVM();

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _vm.OpenRepoCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);

            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void MergeStatusSuccess_MessageIsShown()
        {
            //arrange
            const string sourceBranch = "dev";
            const string destinationBranch = "master";

            var expectedTitle = RubberduckUI.SourceControl_MergeStatus;
            var expectedMessage = string.Format(RubberduckUI.SourceControl_SuccessfulMerge, sourceBranch, destinationBranch);

            SetupValidVbProject();
            SetupVM();

            _vm.Provider = _provider.Object;

            _branchesVM.SourceBranch = sourceBranch;
            _branchesVM.DestinationBranch = destinationBranch;

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _branchesVM.MergeBranchesOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);

            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void ActionFailed_DismissingHidesMessage()
        {
            //arrange
            SetupVM();
            _vm.DisplayErrorMessageGrid = true;

            //act
            _vm.DismissErrorMessageCommand.Execute(null);

            //assert
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenNotAuthorized_LoginIsShown()
        {
            //arrange
            SetupVM();

            const string expectedTitle = "Push Failed.";
            const string expectedMessage = "Request failed with status code: 401";

            _provider.Setup(p => p.Pull())
                .Throws(
                    new SourceControlException(expectedTitle,
                        new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupVM();
            _vm.Provider = _provider.Object;

            //act
            _unsyncedVM.PullCommitsCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayLoginGrid);
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterLogin_NewPresenterIsCreatedWithCredentials()
        {
            //arrange
            const string username = "username";
            var password = new SecureString();
            foreach (var c in "password")
            {
                password.AppendChar(c);
            }

            _configService.Setup(c => c.Create()).Returns(GetDummyConfig());

            SetupValidVbProject();
            SetupVM();

            _vm.Provider = _provider.Object;

            //act
            _vm.CreateProviderWithCredentials(new SecureCredentials(username, password));

            //assert
            _providerFactory.Verify(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<SecureCredentials>(), It.IsAny<ICodePaneWrapperFactory>()));
        }

        [TestMethod]
        public void Login_HideGridOnCancel()
        {
            //arrange
            SetupVM();

            //act
            _vm.LoginGridCancelCommand.Execute(null);

            //Assert
            Assert.IsFalse(_vm.DisplayLoginGrid);
        }

        [TestMethod]
        public void CloneRepo_DisplaysGrid()
        {
            //arrange
            SetupVM();
            _vm.Provider = _provider.Object;

            //act
            _vm.CloneRepoCommand.Execute(null);

            //Assert
            Assert.IsTrue(_vm.DisplayCloneRepoGrid);
        }

        [TestMethod]
        public void CloneRepo_ClonesRepo()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var localDirectory = "C:\\users\\me\\desktop\\git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.CloneRemotePath = remotePath;
            _vm.LocalDirectory = localDirectory;

            //act
            _vm.CloneRepoOkButtonCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.Clone(remotePath, localDirectory));
        }

        [TestMethod]
        public void CloneRepo_HideGridOnClone()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var localDirectory = "C:\\users\\me\\desktop\\git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.CloneRemotePath = remotePath;
            _vm.LocalDirectory = localDirectory;

            //act
            _vm.CloneRepoOkButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(_vm.DisplayCloneRepoGrid);
        }

        [TestMethod]
        public void CloneRepo_HideGridOnCancel()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var localDirectory = "C:\\users\\me\\desktop\\git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.CloneRemotePath = remotePath;
            _vm.LocalDirectory = localDirectory;

            //act
            _vm.CloneRepoCancelButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(_vm.DisplayCloneRepoGrid);
        }

        [TestMethod]
        public void CloneRepo_ClearsRemoteOnClone()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var localDirectory = "C:\\users\\me\\desktop\\git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.CloneRemotePath = remotePath;
            _vm.LocalDirectory = localDirectory;

            //act
            _vm.CloneRepoOkButtonCommand.Execute(null);

            //Assert
            Assert.AreEqual(string.Empty, _vm.CloneRemotePath);
        }

        [TestMethod]
        public void CloneRepo_ClearsRemoteOnClose()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var localDirectory = "C:\\users\\me\\desktop\\git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.CloneRemotePath = remotePath;
            _vm.LocalDirectory = localDirectory;

            //act
            _vm.CloneRepoCancelButtonCommand.Execute(null);

            //Assert
            Assert.AreEqual(string.Empty, _vm.CloneRemotePath);
        }

        [TestMethod]
        public void CloneRepo_ActionFailedEventIsRaised()
        {
            //arrange
            SetupValidVbProject();
            SetupVM();

            _provider.Setup(p => p.Clone(It.IsAny<string>(), It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            //act
            _vm.CloneRepoOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void Publish_DisplaysGrid()
        {
            //arrange
            SetupVM();
            _vm.Provider = _provider.Object;

            //act
            _vm.PublishRepoCommand.Execute(null);

            //Assert
            Assert.IsTrue(_vm.DisplayPublishRepoGrid);
        }

        [TestMethod]
        public void Publish_AddsOrigin()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";
            var branchName = "master";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.PublishRemotePath = remotePath;

            //act
            _vm.PublishRepoOkButtonCommand.Execute(null);

            //Assert
            _provider.Verify(git => git.AddOrigin(remotePath, branchName));
        }

        [TestMethod]
        public void Publish_HideGridOnClone()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.PublishRemotePath = remotePath;

            //act
            _vm.PublishRepoOkButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(_vm.DisplayPublishRepoGrid);
        }

        [TestMethod]
        public void Publish_HideGridOnCancel()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.PublishRemotePath = remotePath;

            //act
            _vm.PublishRepoCancelButtonCommand.Execute(null);

            //Assert
            Assert.IsFalse(_vm.DisplayPublishRepoGrid);
        }

        [TestMethod]
        public void Publish_ClearsRemoteOnCreate()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.PublishRemotePath = remotePath;

            //act
            _vm.PublishRepoOkButtonCommand.Execute(null);

            //Assert
            Assert.AreEqual(string.Empty, _vm.PublishRemotePath);
        }

        [TestMethod]
        public void Publish_ClearsRemoteOnClose()
        {
            //arrange
            var remotePath = @"https://github.com/Hosch250/RemoveParamsTestProj.git";

            SetupValidVbProject();
            SetupVM();
            _vm.Provider = _provider.Object;

            _vm.PublishRemotePath = remotePath;

            //act
            _vm.PublishRepoCancelButtonCommand.Execute(null);

            //Assert
            Assert.AreEqual(string.Empty, _vm.PublishRemotePath);
        }

        [TestMethod]
        public void Publish_NoOpenRepo_ErrorReported()
        {
            //arrange
            _configService.Setup(c => c.Create()).Returns(new SourceControlSettings());

            SetupValidVbProject();
            SetupVM();

            _provider.Setup(p => p.AddOrigin(It.IsAny<string>(), It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            //act
            _vm.PublishRepoOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void Publish_ActionFailedEventIsRaised()
        {
            //arrange
            SetupValidVbProject();
            SetupVM();

            _provider.Setup(p => p.AddOrigin(It.IsAny<string>(), It.IsAny<string>()))
                .Throws(
                    new SourceControlException("A source control exception was thrown.",
                        new LibGit2Sharp.LibGit2SharpException("With an inner libgit2sharp exception"))
                    );

            _vm.Provider = _provider.Object;

            //act
            _vm.PublishRepoOkButtonCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid, "ActionFailedEvent was not raised.");
        }

        [TestMethod]
        public void OpenDirAssignedToRepo_WhenProviderCreationFails_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Throws(new SourceControlException(expectedTitle,
                    new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupValidVbProject();
            SetupVM();

            //assert-act
            Assert.IsFalse(_vm.DisplayErrorMessageGrid);
            _vm.RefreshCommand.Execute(null);

            //assert
            Assert.IsTrue(_vm.DisplayErrorMessageGrid);

            Assert.AreEqual(expectedTitle, _vm.ErrorTitle);
            Assert.AreEqual(expectedMessage, _vm.ErrorMessage);
        }

        [TestMethod]
        public void OnBrowseDefaultRepoLocation_WhenUserConfirms_LocalDirectoryDoesChanges()
        {
            //arrange
            var newPath = "C:\\test\\test2\\git";

            SetupVM();
            _folderBrowser.SetupProperty(f => f.SelectedPath, newPath);
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.OK);

            //act
            _vm.ShowFilePickerCommand.Execute(null);

            //assert
            Assert.AreEqual(newPath, _vm.LocalDirectory);
        }

        [TestMethod]
        public void OnBrowseDefaultRepoLocation_WhenUserCancels_LocalDirectoryDoesNotChange()
        {
            //arrange
            var newPath = "C:\\test\\test2\\git";
            var originalPath = "C:\\users\\me\\desktop\\git";

            SetupVM();
            _vm.LocalDirectory = originalPath;

            _folderBrowser.SetupProperty(f => f.SelectedPath, newPath);
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.Cancel);

            //act
            _vm.ShowFilePickerCommand.Execute(null);

            //assert
            Assert.AreEqual(originalPath, _vm.LocalDirectory);
        }

        private const string DummyRepoId = "SourceControlTest";

        private SourceControlSettings GetDummyConfig()
        {
            return new SourceControlSettings("username", "username@email.com", string.Empty,
                    new List<Repository> { GetDummyRepo() }, "ps.exe");
        }

        private static Repository GetDummyRepo()
        {
            return new Repository
                       (
                           DummyRepoId,
                           @"C:\Users\Christopher\Documents\SourceControlTest",
                           @"https://github.com/ckuhn203/SourceControlTest.git"
                       );
        }
    }
}
