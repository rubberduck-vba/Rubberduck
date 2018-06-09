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

            using (var pane = e.CodePane)
            using (var module = pane.CodeModule)
            {
                var selection = pane.Selection;
                if (selection.StartColumn < 1) { return false; }
                
                if (!e.IsCommitted && e.Character.ToString() == InputToken)
                {
                    var newCode = e.OldCode.Insert(selection.StartColumn - 1, InputToken + OutputToken);
                    module.ReplaceLine(selection.StartLine, newCode);
                    pane.Selection = new Selection(selection.StartLine, selection.StartColumn + 1);
                    e.Handled = true;
                    e.NewCode = newCode;
                    return true;
                }
                return false;
            }
        }
    }
}
