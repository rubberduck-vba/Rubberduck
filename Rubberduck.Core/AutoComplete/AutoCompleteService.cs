using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : SubclassingWindow, IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly List<IAutoComplete> _autoCompletes;

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
            if (e.Keys == Keys.Delete)
            {
                return;
            }

            var pane = e.CodePane;
            using (var module = pane.CodeModule)
            {
                var qualifiedSelection = module.GetQualifiedSelection();
                var selection = qualifiedSelection.Value.Selection;
                var currentContent = module.GetLines(selection);

                foreach (var autoComplete in _autoCompletes.Where(auto => auto.IsEnabled))
                {
                    if (e.Keys == Keys.Back && selection.StartColumn > 1)
                    {
                        // If caret LHS is the AC input token and RHS is the AC output token, we can remove both.
                        // Substring index is 0-based. Selection from code pane is 1-based.
                        // LHS should be at StartColumn - 2, RHS at StartColumn - 1.
                        var caretLHS = currentContent.Substring(Math.Max(0, selection.StartColumn - 2), 1);
                        var caretRHS = currentContent.Substring(selection.StartColumn - 1, 1);

                        if (caretLHS == autoComplete.InputToken && caretRHS == autoComplete.OutputToken 
                            && !string.IsNullOrEmpty(currentContent.Substring(selection.StartColumn - 1)))
                        {
                            var left = currentContent.Substring(0, selection.StartColumn - 2);
                            var right = currentContent.Substring(selection.StartColumn);
                            module.ReplaceLine(selection.StartLine, left + right);
                            pane.Selection = new Selection(selection.StartLine, selection.StartColumn - right.Length);
                            e.Handled = true;
                            break;
                        }
                    }
                    else
                    {
                        if (autoComplete.Execute(e))
                        {
                            break;
                        }
                    }
                }
            }
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
