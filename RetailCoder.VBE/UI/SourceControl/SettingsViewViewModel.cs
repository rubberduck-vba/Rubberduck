using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Windows.Input;
using Rubberduck.Settings;
using Rubberduck.UI.Command;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class SettingsViewViewModel : ViewModelBase, IControlViewModel
    {
        private readonly IConfigurationService<SourceControlConfiguration> _configService;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private readonly SourceControlConfiguration _config;

        public SettingsViewViewModel(
            IConfigurationService<SourceControlConfiguration> configService,
            IFolderBrowserFactory folderBrowserFactory)
        {
            _configService = configService;
            _folderBrowserFactory = folderBrowserFactory;
            _config = _configService.LoadConfiguration();

            UserName = _config.UserName;
            EmailAddress = _config.EmailAddress;
            DefaultRepositoryLocation = _config.DefaultRepositoryLocation;

            _showFilePickerCommand = new DelegateCommand(_ => ShowFilePicker());
            _cancelSettingsChangesCommand = new DelegateCommand(_ => CancelSettingsChanges());
            _updateSettingsCommand = new DelegateCommand(_ => UpdateSettings());
            _showGitIgnoreCommand = new DelegateCommand(_ => ShowGitIgnore(), _ => Provider != null);
            _showGitAttributesCommand = new DelegateCommand(_ => ShowGitAttributes(), _ => Provider != null);
        }

        public ISourceControlProvider Provider { get; set; }

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

        private void ShowFilePicker()
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser("Default Repository Directory"))
            {
                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    DefaultRepositoryLocation = folderPicker.SelectedPath;
                }
            }
        }

        private void CancelSettingsChanges()
        {
            UserName = _config.UserName;
            EmailAddress = _config.EmailAddress;
            DefaultRepositoryLocation = _config.DefaultRepositoryLocation;
        }

        private void UpdateSettings()
        {
            _config.UserName = UserName;
            _config.EmailAddress = EmailAddress;
            _config.DefaultRepositoryLocation = DefaultRepositoryLocation;

            _configService.SaveConfiguration(_config);
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

        private readonly ICommand _showFilePickerCommand;
        public ICommand ShowFilePickerCommand
        {
            get
            {
                return _showFilePickerCommand;
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
        private void RaiseErrorEvent(string message)
        {
            var handler = ErrorThrown;
            if (handler != null)
            {
                handler(this, new ErrorEventArgs(message));
            }
        }
    }
}