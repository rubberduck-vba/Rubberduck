using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
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
        private readonly List<IAutoComplete> _autoCompletes = new List<IAutoComplete>();
        private readonly List<SelfClosingPair> _selfClosingPairs = new List<SelfClosingPair>
        {
            new SelfClosingPair('(', ')'),
            new SelfClosingPair('"', '"'),
            new SelfClosingPair('[', ']'),
            new SelfClosingPair('{', '}'),
        };

        private readonly SelfClosingPairCompletionService _selfClosingPairCompletion;

        private AutoCompleteSettings _settings;
        private bool _popupShown;

        public AutoCompleteService(IGeneralConfigService configService, SelfClosingPairCompletionService selfClosingPairCompletion)
        {
            _selfClosingPairCompletion = selfClosingPairCompletion;
            _configService = configService;
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
            if (_popupShown || (e.Keys != Keys.None && pSelection.LineCount > 1) || e.Keys == Keys.Delete)
            {
                return;
            }

            var currentContent = module.GetLines(pSelection);

            /* "smart concat" // adds a line continuation when {ENTER} is pressed inside a string literal */
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

            var currentCode = e.CurrentLine;
            var currentSelection = e.CurrentSelection;
            var original = new CodeString(currentCode, new Selection(0, currentSelection.EndColumn - 1), new Selection(pSelection.StartLine, 1));

            if (e.Character != default)
            {
                currentCode += e.Character;
            }

            var prettifier = new CodeStringPrettifier(module);
            foreach (var selfClosingPair in _selfClosingPairs)
            {
                CodeString result;
                if (e.Keys == Keys.Back && pSelection.StartColumn > 1)
                {
                    result = _selfClosingPairCompletion.Execute(selfClosingPair, original, e.Keys);
                }
                else
                {
                    result = _selfClosingPairCompletion.Execute(selfClosingPair, original, e.Character, prettifier);
                }

                if (result != default)
                {
                    using (var pane = module.CodePane)
                    {
                        module.DeleteLines(result.SnippetPosition);
                        module.InsertLines(result.SnippetPosition.StartLine, result.Code);
                        pane.Selection = result.SnippetPosition.Offset(result.CaretPosition);
                        e.Handled = true;
                        return;
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
