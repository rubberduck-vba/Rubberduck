using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBase : IAutoComplete
    {
        protected AutoCompleteBase(string inputToken, string outputToken)
        {
            InputToken = inputToken;
            OutputToken = outputToken;
        }

        public bool IsEnabled { get; set; }
        public string InputToken { get; }
        public string OutputToken { get; }

        private readonly Stack<Tuple<int, string>> OriginalLines;

        public void Revert()
        {
            if (OriginalLines.Count > 0)
            {
                var original = OriginalLines.Pop();
                var line = original.Item1;
                var content = original.Item2;

            }
        }

        public virtual bool Execute(AutoCompleteEventArgs e)
        {
            if (!e.IsCharacter)
            {
                return false;
            }

            var module = e.CodeModule;
            using (var pane = module.CodePane)
            {
                var selection = pane.Selection;
                if (e.Character.ToString() == InputToken)
                {
                    var code = module.GetLines(selection).Insert(Math.Max(0, selection.StartColumn - 1), InputToken + OutputToken);
                    module.ReplaceLine(selection.StartLine, code);
                    pane.Selection = new Selection(selection.StartLine, selection.StartColumn + 1);
                    e.Handled = true;
                    return true;
                }
                return false;
            }
        }
    }
}
