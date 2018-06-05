using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly List<IAutoComplete> _autoCompletes;
        private QualifiedSelection? _lastSelection;
        private string _lastCode;
        private string _contentHash;

        public AutoCompleteService(IGeneralConfigService configService, IAutoCompleteProvider provider)
        {
            _configService = configService;
            _autoCompletes = provider.AutoCompletes.ToList();
            UpdateEnabledAutoCompletes(configService.LoadConfiguration());

            _configService.SettingsChanged += ConfigServiceSettingsChanged;
            VBENativeServices.CaretHidden += VBENativeServices_CaretHidden;
        }

        private void ConfigServiceSettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            var config = _configService.LoadConfiguration();
            UpdateEnabledAutoCompletes(config);
        }

        private void UpdateEnabledAutoCompletes(Configuration config)
        {
            foreach (var autoComplete in _autoCompletes)
            {
                foreach (var setting in config.UserSettings.AutoCompleteSettings.AutoCompletes)
                {
                    if (setting.Key == autoComplete.GetType().Name)
                    {
                        autoComplete.IsEnabled = setting.IsEnabled;
                        continue;
                    }
                }
            }
        }

        private void VBENativeServices_CaretHidden(object sender, AutoCompleteEventArgs e)
        {
            if (e.ContentHash == _contentHash || e.OldCode.Length < (_lastCode?.Length ?? 0))
            {
                return;
            }

            var qualifiedSelection = e.CodePane.GetQualifiedSelection();
            var selection = qualifiedSelection.Value.Selection;

            foreach (var autoComplete in _autoCompletes.Where(auto => auto.IsEnabled))
            {
                if (autoComplete.Execute(e))
                {
                    _lastSelection = qualifiedSelection;
                    _lastCode = e.NewCode;
                    using (var module = e.CodePane.CodeModule)
                    {
                        _contentHash = module.ContentHash();
                    }

                    break;
                }
            }
        }

        public void Dispose()
        {
            VBENativeServices.CaretHidden -= VBENativeServices_CaretHidden;
            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }

            _autoCompletes.Clear();
        }
    }
}
