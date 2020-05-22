using Rubberduck.Resources.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public abstract class SettingsViewModelBase<TSettings> : ViewModelBase 
        where TSettings : new()
    {
        protected readonly IConfigurationService<TSettings> Service;

        protected SettingsViewModelBase(IConfigurationService<TSettings> service)
        {
            Service = service;
        }

        public CommandBase ExportButtonCommand { get; protected set; }

        public CommandBase ImportButtonCommand { get; protected set; }

        protected abstract void TransferSettingsToView(TSettings toLoad);
        protected abstract string DialogLoadTitle { get; }
        protected abstract string DialogSaveTitle { get; }

        protected virtual void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = DialogLoadTitle
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName))
                {
                    return;
                }

                Service.Import(dialog.FileName);
                // FIXME transfer settings to view cleaner
                TransferSettingsToView(Service.Read());
            }
        }

        protected virtual void ExportSettings(TSettings settings)
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = DialogSaveTitle
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName))
                {
                    return;
                }

                Service.Export(dialog.FileName);
            }
        }
    }
}
