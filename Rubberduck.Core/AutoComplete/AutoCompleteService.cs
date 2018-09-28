using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IGeneralConfigService _configService;
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
        private bool _enabled = false;
        private bool _initialized;

        public AutoCompleteService(IGeneralConfigService configService, SelfClosingPairCompletionService selfClosingPairCompletion)
        {
            _selfClosingPairCompletion = selfClosingPairCompletion;
            _configService = configService;
            InitializeConfig();
            _configService.SettingsChanged += ConfigServiceSettingsChanged;
        }

        public void Enable()
        {
            if (!_initializing)
            {
                InitializeConfig();
            }

            if (!_enabled)
            {
                VBENativeServices.KeyDown += HandleKeyDown;
                VBENativeServices.IntelliSenseChanged += HandleIntelliSenseChanged;
                _enabled = true;
            }
        }

        private bool _initializing;
        private void InitializeConfig()
        {
            _initializing = true;
            // No reason to think this would throw, but if it does, _initializing state needs to be reset.
            try
            {
                if (!_initialized)
                {
                    var config = _configService.LoadConfiguration();
                    ApplyAutoCompleteSettings(config);
                }
            }
            finally
            {
                _initializing = false;
            }            
        }

        public void Disable()
        {
            if (_enabled && _initialized)
            {
                VBENativeServices.KeyDown -= HandleKeyDown;
                VBENativeServices.IntelliSenseChanged -= HandleIntelliSenseChanged;
                _enabled = false;
            }
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
            if (_settings.IsEnabled)
            {
                Enable();
            }
            else
            {
                Disable();
            }
            _initialized = true;
        }

        private void HandleKeyDown(object sender, AutoCompleteEventArgs e)
        {
            if (e.Character == default && !e.IsDelete)
            {
                return;
            }

            var module = e.CodeModule;
            var qualifiedSelection = module.GetQualifiedSelection();
            Debug.Assert(qualifiedSelection != null, nameof(qualifiedSelection) + " != null");
            var pSelection = qualifiedSelection.Value.Selection;

            if (_popupShown || pSelection.LineCount > 1 || e.IsDelete)
            {
                return;
            }

            var currentContent = module.GetLines(pSelection);
            if (HandleSmartConcat(e, pSelection, currentContent, module))
            {
                return;
            }

            HandleSelfClosingPairs(e, module, pSelection);
        }

        private void HandleSelfClosingPairs(AutoCompleteEventArgs e, ICodeModule module, Selection pSelection)
        {
            if (!pSelection.IsSingleCharacter)
            {
                return;
            }

            var currentCode = e.CurrentLine;
            var currentSelection = e.CurrentSelection;
            //var surroundingCode = GetSurroundingCode(module, currentSelection); // todo: find a way to parse the current instruction

            var original = new CodeString(currentCode, new Selection(0, currentSelection.EndColumn - 1), new Selection(pSelection.StartLine, 1));

            var prettifier = new CodeStringPrettifier(module);
            foreach (var selfClosingPair in _selfClosingPairs)
            {
                CodeString result;
                if (e.Character == '\b' && pSelection.StartColumn > 1)
                {
                    result = _selfClosingPairCompletion.Execute(selfClosingPair, original, '\b');
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

        /// <summary>
        /// Adds a line continuation when {ENTER} is pressed inside a string literal; returns false otherwise.
        /// </summary>
        private bool HandleSmartConcat(AutoCompleteEventArgs e, Selection pSelection, string currentContent, ICodeModule module)
        {
            var shouldHandle = _settings.EnableSmartConcat &&
                               e.Character == '\r' &&
                               IsInsideStringLiteral(pSelection, ref currentContent);

            var lastIndexLeftOfCaret = currentContent.Length > 2 ? currentContent.Substring(0, pSelection.StartColumn - 1).LastIndexOf('"') : 0;
            if (shouldHandle && lastIndexLeftOfCaret > 0)
            {
                var indent = currentContent.NthIndexOf('"', 1);
                var whitespace = new string(' ', indent);
                var code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";

                if (e.ControlDown)
                {
                    code = $"{currentContent.Substring(0, pSelection.StartColumn - 1)}\" & vbNewLine & _\r\n{whitespace}\"{currentContent.Substring(pSelection.StartColumn - 1)}";

                }

                module.ReplaceLine(pSelection.StartLine, code);
                using (var pane = module.CodePane)
                {
                    pane.Selection = new Selection(pSelection.StartLine + 1, indent + currentContent.Substring(pSelection.StartColumn - 2).Length);
                    e.Handled = true;
                    return true;
                }
            }

            return false;
        }

        private string GetSurroundingCode(ICodeModule module, Selection selection)
        {
            // throws AccessViolationException!
            var declarationLines = module.CountOfDeclarationLines;
            if (selection.StartLine <= declarationLines)
            {
                return module.GetLines(1, declarationLines);
            }

            var currentProc = module.GetProcOfLine(selection.StartLine);
            var procKind = module.GetProcKindOfLine(selection.StartLine);
            var procStart = module.GetProcStartLine(currentProc, procKind);
            var lineCount = module.GetProcCountLines(currentProc, procKind);
            return module.GetLines(procStart, lineCount);
        }

        private bool IsInsideStringLiteral(Selection pSelection, ref string currentContent)
        {
            if (!currentContent.Substring(pSelection.StartColumn - 1).Contains("\"") || 
                currentContent.StripStringLiterals().HasComment(out _))
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
            Disable();
            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
            }
        }
    }
}
