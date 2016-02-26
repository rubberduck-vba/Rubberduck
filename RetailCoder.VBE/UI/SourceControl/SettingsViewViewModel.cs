using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.SourceControl
{
    public class SettingsViewViewModel : ViewModelBase
    {
        public SettingsViewViewModel()
        {
            _showFilePickerCommand = new DelegateCommand(_ => ShowFilePicker());
            _cancelSettingsChangesCommand = new DelegateCommand(_ => CancelSettingsChanges());
            _updateSettingsCommand = new DelegateCommand(_ => UpdateSettings());
            _showGitIgnoreCommand = new DelegateCommand(_ => ShowGitIgnore());
            _showGitAttributesCommand = new DelegateCommand(_ => ShowGitAttributes());
        }

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

        private string _defaultRepoLocation;
        public string DefaultRepoLocation
        {
            get { return _defaultRepoLocation; }
            set
            {
                if (_defaultRepoLocation != value)
                {
                    _defaultRepoLocation = value;
                    OnPropertyChanged();
                }
            }
        }

        private void ShowFilePicker()
        {
        }

        private void CancelSettingsChanges()
        {
        }

        private void UpdateSettings()
        {
        }

        private void ShowGitIgnore()
        {
        }

        private void ShowGitAttributes()
        {
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
    }
}