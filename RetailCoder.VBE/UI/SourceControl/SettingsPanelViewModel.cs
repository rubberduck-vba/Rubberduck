using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using NLog;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class SettingsPanelViewModel : ViewModelBase, IControlViewModel, IDisposable
    {
        private readonly IConfigProvider<SourceControlSettings> _configService;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly IOpenFileDialog _openFileDialog;
        private readonly SourceControlSettings _config;

        public SettingsPanelViewModel(
            IConfigProvider<SourceControlSettings> configService,
            IFolderBrowserFactory folderBrowserFactory,
            IOpenFileDialog openFileDialog)
        {
            _configService = configService;
            _folderBrowserFactory = folderBrowserFactory;
            _config = _configService.Create();

            _openFileDialog = openFileDialog;
            _openFileDialog.Filter = "Executables (*.exe)|*.exe|All files (*.*)|*.*";
            _openFileDialog.Multiselect = false;
            _openFileDialog.ReadOnlyChecked = true;
            _openFileDialog.CheckFileExists = true;

            UserName = _config.UserName;
            EmailAddress = _config.EmailAddress;
            DefaultRepositoryLocation = _config.DefaultRepositoryLocation;
            CommandPromptLocation = _config.CommandPromptLocation;

            ShowDefaultRepoFolderPickerCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowDefaultRepoFolderPicker());
            ShowCommandPromptExePickerCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowCommandPromptExePicker());
            CancelSettingsChangesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CancelSettingsChanges());
            UpdateSettingsCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => UpdateSettings());
            ShowGitIgnoreCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowGitIgnore(), _ => Provider != null);
            ShowGitAttributesCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowGitAttributes(), _ => Provider != null);
        }

        public ISourceControlProvider Provider { get; set; }
        public void RefreshView() { } // nothing to refresh here

        public void ResetView()
        {
            Provider = null;
        }

        public SourceControlTab Tab => SourceControlTab.Settings;

        private string _userName;
        public string UserName
        {
            get => _userName;
            set
            {
                if (_userName != value)
                {
                    _userName = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _emailAddress;
        public string EmailAddress
        {
            get => _emailAddress;
            set
            {
                if (_emailAddress != value)
                {
                    _emailAddress = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _defaultRepositoryLocation;
        public string DefaultRepositoryLocation
        {
            get => _defaultRepositoryLocation;
            set
            {
                if (_defaultRepositoryLocation != value)
                {
                    _defaultRepositoryLocation = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _commandPromptExeLocation;
        public string CommandPromptLocation
        {
            get => _commandPromptExeLocation;
            set
            {
                if (_commandPromptExeLocation != value)
                {
                    _commandPromptExeLocation = value;
                    OnPropertyChanged();
                }
            }
        }

        private void ShowDefaultRepoFolderPicker()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser(RubberduckUI.SourceControl_FilePickerDefaultRepoHeader))
            {
                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    DefaultRepositoryLocation = folderPicker.SelectedPath;
                }
            }
        }

        private void ShowCommandPromptExePicker()
        {
            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                CommandPromptLocation = _openFileDialog.FileName;
            }
        }

        private void CancelSettingsChanges()
        {
            UserName = _config.UserName;
            EmailAddress = _config.EmailAddress;
            DefaultRepositoryLocation = _config.DefaultRepositoryLocation;
            CommandPromptLocation = _config.CommandPromptLocation;
        }

        private void UpdateSettings()
        {
            _config.UserName = UserName;
            _config.EmailAddress = EmailAddress;
            _config.DefaultRepositoryLocation = DefaultRepositoryLocation;
            _config.CommandPromptLocation = CommandPromptLocation;

            _configService.Save(_config);

            RaiseErrorEvent(RubberduckUI.SourceControl_UpdateSettingsTitle,
                RubberduckUI.SourceControl_UpdateSettingsMessage, NotificationType.Info);
        }

        private void ShowGitIgnore()
        {
            OpenFileInExternalEditor(GitSettingsFile.Ignore);
        }

        private void ShowGitAttributes()
        {
            OpenFileInExternalEditor(GitSettingsFile.Attributes);
        }

        private void OpenFileInExternalEditor(GitSettingsFile fileType)
        {
            var fileName = string.Empty;
            var defaultContents = string.Empty;
            switch (fileType)
            {
                case GitSettingsFile.Ignore:
                    fileName = ".gitignore";
                    defaultContents = DefaultSettings.GitIgnoreText();
                    break;
                case GitSettingsFile.Attributes:
                    fileName = ".gitattributes";
                    defaultContents = DefaultSettings.GitAttributesText();
                    break;
            }

            var repo = Provider.CurrentRepository;
            var filePath = Path.Combine(repo.LocalLocation, fileName);

            if (!File.Exists(filePath))
            {
                File.WriteAllText(filePath, defaultContents);
            }

            Process.Start(filePath);
        }

        public CommandBase ShowDefaultRepoFolderPickerCommand { get; }

        public CommandBase ShowCommandPromptExePickerCommand { get; }

        public CommandBase CancelSettingsChangesCommand { get; }

        public CommandBase UpdateSettingsCommand { get; }

        public CommandBase ShowGitIgnoreCommand { get; }

        public CommandBase ShowGitAttributesCommand { get; }

        public event EventHandler<ErrorEventArgs> ErrorThrown;

        private void RaiseErrorEvent(string message, Exception innerException, NotificationType notificationType)
        {
            ErrorThrown?.Invoke(this, new ErrorEventArgs(message, innerException, notificationType));
        }

        private void RaiseErrorEvent(string title, string message, NotificationType notificationType)
        {
            ErrorThrown?.Invoke(this, new ErrorEventArgs(title, message, notificationType));
        }

        public void Dispose()
        {
            _openFileDialog?.Dispose();
        }
    }
}