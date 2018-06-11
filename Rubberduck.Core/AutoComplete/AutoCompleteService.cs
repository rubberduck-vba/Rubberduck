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

            var module = e.CodeModule;
            var qualifiedSelection = module.GetQualifiedSelection();
            var selection = qualifiedSelection.Value.Selection;
            var currentContent = module.GetLines(selection);

            foreach (var autoComplete in _autoCompletes.Where(auto => auto.IsEnabled))
            {
                if (e.Keys == Keys.Back && selection.StartColumn > 1)
                {
                    if (HandleBackspace(e, autoComplete))
                    {
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

        private bool HandleBackspace(AutoCompleteEventArgs e, IAutoComplete autoComplete)
        {
            var isInlineAutoComplete = autoComplete.InputToken.Length == 1;
            if (isInlineAutoComplete)
            {
                var code = e.CurrentLine;
                // If caret LHS is the AC input token and RHS is the AC output token, we can remove both.
                // Substring index is 0-based. Selection from code pane is 1-based.
                // LHS should be at StartColumn - 2, RHS at StartColumn - 1.
                var caretLHS = code.Substring(Math.Max(0, e.CurrentSelection.StartColumn - 2), 1);
                var caretRHS = code.Length >= e.CurrentSelection.StartColumn
                    ? code.Substring(e.CurrentSelection.StartColumn - 1, 1)
                    : string.Empty;

                if (caretLHS == autoComplete.InputToken && caretRHS == autoComplete.OutputToken
                    /*&& !string.IsNullOrEmpty(code.Substring(e.CurrentSelection.StartColumn - 1))*/)
                {
                    var left = code.Substring(0, e.CurrentSelection.StartColumn - 2);
                    var right = code.Substring(e.CurrentSelection.StartColumn);
                    using (var pane = e.CodeModule.CodePane)
                    {
                        e.CodeModule.ReplaceLine(e.CurrentSelection.StartLine, left + right);
                        pane.Selection = new Selection(e.CurrentSelection.StartLine, e.CurrentSelection.StartColumn - 1);
                        e.Handled = true;
                    }
                    return true;
                }
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
