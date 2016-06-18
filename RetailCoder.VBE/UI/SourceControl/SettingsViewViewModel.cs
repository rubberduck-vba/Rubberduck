using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Windows.Input;
using NLog;
using Rubberduck.UI.Command;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class SettingsViewViewModel : ViewModelBase, IControlViewModel, IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly ISourceControlConfigProvider _configService;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly IOpenFileDialog _openFileDialog;
        private readonly SourceControlSettings _config;

        public SettingsViewViewModel(
            ISourceControlConfigProvider configService,
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

            _showDefaultRepoFolderPickerCommand = new DelegateCommand(_ => ShowDefaultRepoFolderPicker());
            _showCommandPromptExePickerCommand = new DelegateCommand(_ => ShowCommandPromptExePicker());
            _cancelSettingsChangesCommand = new DelegateCommand(_ => CancelSettingsChanges());
            _updateSettingsCommand = new DelegateCommand(_ => UpdateSettings());
            _showGitIgnoreCommand = new DelegateCommand(_ => ShowGitIgnore(), _ => Provider != null);
            _showGitAttributesCommand = new DelegateCommand(_ => ShowGitAttributes(), _ => Provider != null);
        }

        public ISourceControlProvider Provider { get; set; }
        public void RefreshView() {} // nothing to refresh here

        public SourceControlTab Tab { get { return SourceControlTab.Settings; } }

        private string _userName;
        public string UserName
        {
            get { return _userName; }
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
            get { return _emailAddress; }
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
            get { return _defaultRepositoryLocation; }
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
            get { return _commandPromptExeLocation; }
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
            Logger.Trace("Settings changes canceled");

            UserName = _config.UserName;
            EmailAddress = _config.EmailAddress;
            DefaultRepositoryLocation = _config.DefaultRepositoryLocation;
            CommandPromptLocation = _config.CommandPromptLocation;
        }

        private void UpdateSettings()
        {
            Logger.Trace("Settings changes saved");

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

        private readonly ICommand _showDefaultRepoFolderPickerCommand;
        public ICommand ShowDefaultRepoFolderPickerCommand
        {
            get
            {
                return _showDefaultRepoFolderPickerCommand;
            }
        }

        private readonly ICommand _showCommandPromptExePickerCommand;
        public ICommand ShowCommandPromptExePickerCommand
        {
            get
            {
                return _showCommandPromptExePickerCommand;
            }
        }

        private readonly ICommand _cancelSettingsChangesCommand;
        public ICommand CancelSettingsChangesCommand
        {
            get
            {
                return _cancelSettingsChangesCommand;
            }
        }

        private readonly ICommand _updateSettingsCommand;
        public ICommand UpdateSettingsCommand
        {
            get
            {
                return _updateSettingsCommand;
            }
        }

        private readonly ICommand _showGitIgnoreCommand;
        public ICommand ShowGitIgnoreCommand
        {
            get
            {
                return _showGitIgnoreCommand;
            }
        }

        private readonly ICommand _showGitAttributesCommand;
        public ICommand ShowGitAttributesCommand
        {
            get
            {
                return _showGitAttributesCommand;
            }
        }

        public event EventHandler<ErrorEventArgs> ErrorThrown;
        private void RaiseErrorEvent(string message, string innerMessage, NotificationType notificationType)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message, innerMessage, notificationType));
            }
        }

        public void Dispose()
        {
            if (_openFileDialog != null)
            {
                _openFileDialog.Dispose();
            }
        }
    }
}