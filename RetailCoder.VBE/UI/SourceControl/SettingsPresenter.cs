using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.SourceControl;
using Rubberduck.Settings;

namespace Rubberduck.UI.SourceControl
{
    public interface ISettingsPresenter : IProviderPresenter
    {
        void RefreshView();
    }

    public class SettingsPresenter : ISettingsPresenter
    {
        private readonly ISettingsView _view;
        private readonly IConfigurationService<SourceControlConfiguration> _configurationService;
        private readonly IFolderBrowserFactory _folderBrowserFactory;
        private SourceControlConfiguration _config;

        public ISourceControlProvider Provider{ get; set; }

        public SettingsPresenter(ISettingsView view, IConfigurationService<SourceControlConfiguration> configService, IFolderBrowserFactory folderBrowserFactory)
        {
            _configurationService = configService;
            _folderBrowserFactory = folderBrowserFactory;

            _view = view;

            _view.BrowseDefaultRepositoryLocation += OnBrowseDefaultRepositoryLocation;
            _view.Cancel += OnCancel;
            _view.EditIgnoreFile += OnEditIgnoreFile;
            _view.EditAttributesFile += OnEditAttributesFile;
            _view.Save += OnSave;
        }

        public void RefreshView()
        {
            _config = _configurationService.LoadConfiguration();

            _view.UserName = _config.UserName;
            _view.EmailAddress = _config.EmailAddress;
            _view.DefaultRepositoryLocation = _config.DefaultRepositoryLocation;
        }

        private void OnSave(object sender, EventArgs e)
        {
            if (_config == null)
            {
                _config = _configurationService.LoadConfiguration();
            }

            _config.EmailAddress = _view.EmailAddress;
            _config.UserName = _view.UserName;
            _config.DefaultRepositoryLocation = _view.DefaultRepositoryLocation;

            _configurationService.SaveConfiguration(_config);
        }

        private void OnEditAttributesFile(object sender, EventArgs e)
        {
            OpenFileInExternalEditor(".gitattributes");
        }

        private void OnEditIgnoreFile(object sender, EventArgs e)
        {
            OpenFileInExternalEditor(".gitignore");
        }

        private void OpenFileInExternalEditor(string fileName)
        {
            if (this.Provider == null)
            {
                return;
            }

            var repo = this.Provider.CurrentRepository;
            var filePath = System.IO.Path.Combine(repo.LocalLocation, fileName);

            if (System.IO.File.Exists(filePath))
            {
                Process.Start(filePath);
            }
        }

        private void OnCancel(object sender, EventArgs e)
        {
            RefreshView();
        }

        private void OnBrowseDefaultRepositoryLocation(object sender, EventArgs e)
        {
            using (var folderPicker = _folderBrowserFactory.CreateFolderBrowser("Default Repository Directory"))
            {
                if (folderPicker.ShowDialog() == DialogResult.OK)
                {
                    _view.DefaultRepositoryLocation = folderPicker.SelectedPath;
                }
            }
        }
    }
}
