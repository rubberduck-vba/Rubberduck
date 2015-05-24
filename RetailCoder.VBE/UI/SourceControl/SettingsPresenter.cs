using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using Rubberduck.Config;

namespace Rubberduck.UI.SourceControl
{
    public class SettingsPresenter : IProviderPresenter
    {
        private readonly ISettingsView _view;
        private readonly IConfigurationService<SourceControlConfiguration> _configurationService;
        private SourceControlConfiguration _config;

        public ISourceControlProvider Provider{ get; set; }

        public SettingsPresenter(ISettingsView view, IConfigurationService<SourceControlConfiguration> configService )
        {
            _configurationService = configService;

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
            throw new NotImplementedException();
        }

        private void OnEditIgnoreFile(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void OnCancel(object sender, EventArgs e)
        {
            RefreshView();
        }

        private void OnBrowseDefaultRepositoryLocation(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
