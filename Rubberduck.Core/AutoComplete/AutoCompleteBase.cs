using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete
{

    public abstract class AutoCompleteBase : IAutoComplete
    {
        protected AutoCompleteBase(string inputToken, string outputToken)
        {
            InputToken = inputToken;
            OutputToken = outputToken;
        }

        public bool IsInlineCharCompletion => InputToken.Length == 1 && OutputToken.Length == 1;
        public bool IsEnabled { get; set; }
        public string InputToken { get; }
        public string OutputToken { get; }

        public virtual bool Execute(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            var input = e.Character.ToString();
            if (!IsMatch(input))
            {
                return false;
            }

            var module = e.CodeModule;
            using (var pane = module.CodePane)
            {
                var pSelection = pane.Selection;
                var zSelection = pSelection.ToZeroBased();

                var original = module.GetLines(pSelection);
                var nextChar = zSelection.StartColumn == original.Length ? string.Empty : original.Substring(zSelection.StartColumn, 1);
                if (input == InputToken && (input != OutputToken || nextChar != OutputToken))
                {
                    string code;
                    if (!StripExplicitCallStatement(ref original, ref pSelection))
                    {
                        code = original.Insert(Math.Max(0, zSelection.StartColumn), InputToken + OutputToken);
                    }
                    else
                    {
                        code = original;
                    }
                    module.ReplaceLine(pSelection.StartLine, code);

                    var newCode = module.GetLines(pSelection);
                    if (newCode.Equals(code, StringComparison.OrdinalIgnoreCase))
                    {
                        pane.Selection = new Selection(pSelection.StartLine, pSelection.StartColumn + 1);
                    }
                    else
                    {
                        // VBE added a space; need to compensate:
                        pane.Selection = new Selection(pSelection.StartLine, GetPrettifiedCaretPosition(pSelection, code, newCode));
                    }
                    e.Handled = true;
                    return true;
                }
                else if (input == OutputToken && nextChar == OutputToken)
                {
                    // just move caret one character to the right & suppress the keypress
                    pane.Selection = new Selection(pSelection.StartLine, GetPrettifiedCaretPosition(pSelection, original, original) + 1);
                    e.Handled = true;
                    return true;
                }
                return false;
            }
        }

        private bool StripExplicitCallStatement(ref string code, ref Selection pSelection)
        {
            // VBE will "helpfully" strip empty parentheses in 'Call Something()'
            // ...and there's no way around it. since Call statement is optional and obsolete,
            // this function strips it
            var pattern = @"\bCall\b\s+";
            if (Regex.IsMatch(code, pattern, RegexOptions.IgnoreCase))
            {
                pSelection = new Selection(pSelection.StartLine, pSelection.StartColumn - "Call ".Length);
                code = Regex.Replace(code, pattern, string.Empty, RegexOptions.IgnoreCase);
                return true;
            }
            return false;
        }

        private int GetPrettifiedCaretPosition(Selection pSelection, string insertedCode, string prettifiedCode)
        {
            var zSelection = pSelection.ToZeroBased();

            var outputTokenIndices = new List<int>();
            for (int i = 0; i < insertedCode.Length; i++)
            {
                var character = insertedCode[i].ToString();
                if (character == OutputToken)
                {
                    outputTokenIndices.Add(i);
                }
            }
            if (!outputTokenIndices.Any())
            {
                return pSelection.EndColumn;
            }
            var firstAfterCaret = outputTokenIndices.Where(i => i > zSelection.StartColumn).Min();

            var prettifiedTokenIndices = new List<int>();
            for (int i = 0; i < prettifiedCode.Length; i++)
            {
                var character = prettifiedCode[i].ToString();
                if (character == OutputToken)
                {
                    prettifiedTokenIndices.Add(i);
                }
            }

            return prettifiedTokenIndices.Any()
                ? prettifiedTokenIndices[outputTokenIndices.IndexOf(firstAfterCaret)] + 1
                : prettifiedCode.Length + 2;
        }

        public virtual bool IsMatch(string input) => 
            (IsInlineCharCompletion && !string.IsNullOrEmpty(input) && (input == InputToken || input == OutputToken));
    }
}
