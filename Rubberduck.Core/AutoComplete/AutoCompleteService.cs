using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : SubclassingWindow, IDisposable
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
            VBENativeServices.KeyDown += HandleKeyDown;
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
                var setting = config.UserSettings.AutoCompleteSettings.AutoCompletes.FirstOrDefault(s => s.Key == autoComplete.GetType().Name);
                if (setting != null && autoComplete.IsEnabled != setting.IsEnabled)
                {
                    autoComplete.IsEnabled = setting.IsEnabled;
                    continue;
                }
            }
        }

        private void HandleKeyDown(object sender, AutoCompleteEventArgs e)
        {
            if (e.ContentHash == _contentHash)
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

        /// <summary>
        /// Handles a WM.KeyDown event, before the key is written to the code pane.
        /// Return <c>true</c> to "swallow" the key.
        /// </summary>
        public bool HandleKeyPress(ICodeModule module, Keys keys)
        {
            if (module.ContentHash() == _contentHash)
            {
                return false;
            }

            var selection = module.GetQualifiedSelection().Value.Selection;

            if (keys == Keys.Back)
            {
                // if cursor LHS is opening and RHS is closing any inline autocomplete, delete the next character.
                return true;
            }

            return false;
        }

        public void Dispose()
        {
            VBENativeServices.KeyDown -= HandleKeyDown;
            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }

            _autoCompletes.Clear();
        }
    }
}
