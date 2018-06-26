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
                    var code = original.Insert(Math.Max(0, zSelection.StartColumn), InputToken + OutputToken);
                    module.ReplaceLine(pSelection.StartLine, code);

                    var newCode = module.GetLines(pSelection);
                    if (newCode == code)
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
                    pane.Selection = new Selection(pSelection.StartLine, pSelection.StartColumn + 2);
                    e.Handled = true;
                    return true;
                }
                return false;
            }
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

            return prettifiedTokenIndices[outputTokenIndices.IndexOf(firstAfterCaret)] + 1;
        }

        public virtual bool IsMatch(string input) => 
            (IsInlineCharCompletion && !string.IsNullOrEmpty(input) && (input == InputToken || input == OutputToken));
    }
}
