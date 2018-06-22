using Rubberduck.Settings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using System;

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
                var selection = pane.Selection;
                var original = module.GetLines(selection);
                var nextChar = selection.StartColumn - 1 == original.Length ? string.Empty : original.Substring(selection.StartColumn - 1, 1);
                if (input == InputToken && (input != OutputToken || nextChar != OutputToken))
                {
                    var code = original.Insert(Math.Max(0, selection.StartColumn - 1), InputToken + OutputToken);
                    module.ReplaceLine(selection.StartLine, code);
                    var newCode = module.GetLines(selection);
                    if (newCode == code)
                    {
                        pane.Selection = new Selection(selection.StartLine, selection.StartColumn + 1);
                    }
                    else
                    {
                        // VBE added a space; need to compensate:
                        pane.Selection = new Selection(selection.StartLine, selection.StartColumn + 2);
                    }
                    e.Handled = true;
                    return true;
                }
                else if (input == OutputToken && nextChar == OutputToken)
                {
                    // just move caret one character to the right & suppress the keypress
                    pane.Selection = new Selection(selection.StartLine, selection.StartColumn + 2);
                    e.Handled = true;
                    return true;
                }
                return false;
            }
        }

        public virtual bool IsMatch(string input) => 
            (IsInlineCharCompletion && !string.IsNullOrEmpty(input) && (input == InputToken || input == OutputToken));
    }
}
