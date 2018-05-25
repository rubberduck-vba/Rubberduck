using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBase : IAutoComplete
    {
        protected AutoCompleteBase(string inputToken, string outputToken)
        {
            InputToken = inputToken;
            OutputToken = outputToken;
        }

        public bool IsEnabled => true;
        public string InputToken { get; }
        public string OutputToken { get; }

        public virtual bool Execute(AutoCompleteEventArgs e)
        {
            var selection = e.CodePane.Selection;
            if (selection.StartColumn < 2) { return false; }

            if (!e.IsCommitted && e.OldCode.Substring(selection.StartColumn - 2, 1) == InputToken)
            {
                using (var module = e.CodePane.CodeModule)
                {
                    var newCode = e.OldCode.Insert(selection.StartColumn - 1, OutputToken);
                    module.ReplaceLine(e.CodePane.Selection.StartLine, newCode);
                    e.CodePane.Selection = selection;
                    e.NewCode = newCode;
                    return true;
                }
            }
            return false;
        }
    }
}
