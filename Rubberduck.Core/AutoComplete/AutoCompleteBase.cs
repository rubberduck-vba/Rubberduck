using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBase : IAutoComplete
    {
        public bool IsEnabled => true;
        public abstract string InputToken { get; }
        public abstract string OutputToken { get; }

        public virtual bool Execute(AutoCompleteEventArgs e)
        {
            var selection = e.CodePane.Selection;
            if (selection.StartColumn < 2) { return false; }

            if (!e.IsCommitted && e.Code.Substring(selection.StartColumn - 2, 1) == InputToken)
            {
                using (var module = e.CodePane.CodeModule)
                {
                    var replacement = e.Code.Insert(selection.StartColumn - 1, OutputToken);
                    module.ReplaceLine(e.CodePane.Selection.StartLine, replacement);
                    e.CodePane.Selection = selection;
                    e.ReplacementLineContent = replacement;
                    return true;
                }
            }
            return false;
        }
    }
}
