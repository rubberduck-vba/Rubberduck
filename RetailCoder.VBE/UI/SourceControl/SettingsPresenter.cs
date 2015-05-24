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
        private IConfigurationService<SourceControlConfiguration> _configurationService;

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

        private void OnSave(object sender, EventArgs e)
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        private void OnBrowseDefaultRepositoryLocation(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
