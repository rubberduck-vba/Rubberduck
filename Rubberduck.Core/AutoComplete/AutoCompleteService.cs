using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.BlockCompletion;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.WindowsApi;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IGeneralConfigService _configService;
        private readonly List<IAutoComplete> _autoCompletes;

        private readonly BlockCompletionService _blockCompletion = new BlockCompletionService();

        private AutoCompleteSettings _settings;
        private bool _popupShown;

        public AutoCompleteService(IGeneralConfigService configService, IAutoCompleteProvider provider)
        {
            _configService = configService;
            _autoCompletes = provider.AutoCompletes.ToList();

            _configService.SettingsChanged += ConfigServiceSettingsChanged;
            VBENativeServices.KeyDown += HandleKeyDown;
            VBENativeServices.IntelliSenseChanged += HandleIntelliSenseChanged;
        }

        private void HandleIntelliSenseChanged(object sender, IntelliSenseEventArgs e)
        {
            _popupShown = e.Visible;
        }

        private void ConfigServiceSettingsChanged(object sender, ConfigurationChangedEventArgs e)
        {
            var config = _configService.LoadConfiguration();
            ApplyAutoCompleteSettings(config);
        }
        
        public void ApplyAutoCompleteSettings(Configuration config)
        {
            _settings = config.UserSettings.AutoCompleteSettings;
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
            if (e.Character == default && e.Keys == Keys.None)
            {
                return;
            }

            var module = e.CodeModule;
            var qualifiedSelection = module.GetQualifiedSelection();
            var pSelection = qualifiedSelection.Value.Selection;

            if (_popupShown || (e.Keys != Keys.None && pSelection.LineCount > 1))
            {
                return;
            }

            var currentContent = module.GetLines(pSelection);
            if (e.Keys == Keys.Enter && _settings.EnableSmartConcat && IsInsideStringLiteral(pSelection, ref currentContent))
            {
                var indent = currentContent.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);
                var code = $"{currentContent} & _\r\n{whitespace}\"";
                module.ReplaceLine(pSelection.StartLine, code);
                using (var pane = module.CodePane)
                {
                    pane.Selection = new Selection(pSelection.StartLine + 1, indent + 2);
                    e.Handled = true;
                    return;
                }
            }

            var handleDelete = e.Keys == Keys.Delete && pSelection.EndColumn <= currentContent.Length;
            var handleBackspace = e.Keys == Keys.Back && pSelection.StartColumn > 1;
            var handleTab = e.Keys == Keys.Tab && !pSelection.IsSingleCharacter;
            var handleEnter = e.Keys == Keys.Enter && !pSelection.IsSingleCharacter;

            var currentCode = e.CurrentLine;
            var currentSelection = e.CurrentSelection;
            if (e.Character != default)
            {
                currentCode += e.Character;
                currentSelection = new Selection(e.CurrentSelection.StartLine, e.CurrentSelection.StartColumn, e.CurrentSelection.EndLine, e.CurrentSelection.EndColumn);
            }

            using (var pane = module.CodePane)
            {
                if (_blockCompletion.Run(e.Keys, currentCode, currentSelection, module, out string newCode, out Selection newSelection))
                {
                    if (newCode.Trim() != e.CurrentLine)
                    {
                        module.ReplaceLine(currentSelection.StartLine, newCode);
                    }
                    if (pane.Selection != newSelection)
                    {
                        pane.Selection = newSelection;
                    }
                    e.Handled = true;
                    return;
                }
            }

            foreach (var autoComplete in _autoCompletes.Where(auto => auto.IsEnabled && auto.IsInlineCharCompletion))
            {
                if ((handleTab || handleEnter) && autoComplete.IsMatch(currentContent))
                {
                    using (var pane = module.CodePane)
                    {
                        if (!string.IsNullOrWhiteSpace(module.GetLines(pSelection.StartLine + 1, 1)))
                        {
                            module.InsertLines(pSelection.StartLine + 1, string.Empty);
                            e.Handled = e.Keys != Keys.Tab; // swallow ENTER, let TAB through
                        }
                        else
                        {
                            pane.Selection = new Selection(pSelection.StartLine + 1, pSelection.EndColumn);
                            e.Handled = true; // base.Execute added the indentation as applicable already.
                        }
                        break;
                    }
                }
                else if (handleBackspace)
                {
                    if (DeleteAroundCaret(e, autoComplete))
                    {
                        break;
                    }
                }
                else
                {
                    if (autoComplete.Execute(e, _settings))
                    {
                        break;
                    }
                }
            }
        }

        private bool IsInsideStringLiteral(Selection pSelection, ref string currentContent)
        {
            if (!currentContent.Contains("\"") || currentContent.StripStringLiterals().HasComment(out _))
            {
                return false;
            }

            var zSelection = pSelection.ToZeroBased();
            var leftOfCaret = currentContent.Substring(0, zSelection.StartColumn);
            var rightOfCaret = currentContent.Substring(Math.Min(zSelection.StartColumn + 1, currentContent.Length - 1));
            if (!rightOfCaret.Contains("\""))
            {
                // the string isn't terminated, but VBE would terminate it here.
                currentContent += "\"";
                rightOfCaret += "\"";
            }

            // odd number of double quotes on either side of the caret means we're inside a string literal, right?
            return (leftOfCaret.Count(c => c.Equals('"')) % 2) != 0 &&
                   (rightOfCaret.Count(c => c.Equals('"')) % 2) != 0;
        }

        private bool DeleteAroundCaret(AutoCompleteEventArgs e, IAutoComplete autoComplete)
        {
            if (autoComplete.IsInlineCharCompletion)
            {
                var code = e.CurrentLine;
                // If caret LHS is the AC input token and RHS is the AC output token, we can remove both.
                // Substring index is 0-based. Selection from code pane is 1-based.
                // LHS should be at StartColumn - 2, RHS at StartColumn - 1.
                var caretLHS = code.Substring(Math.Max(0, e.CurrentSelection.StartColumn - 2), 1);
                var caretRHS = code.Length >= e.CurrentSelection.StartColumn
                    ? code.Substring(e.CurrentSelection.StartColumn - 1, 1)
                    : string.Empty;

                if (caretLHS == autoComplete.InputToken && caretRHS == autoComplete.OutputToken)
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
