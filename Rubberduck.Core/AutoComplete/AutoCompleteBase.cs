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

        public bool IsEnabled { get; set; }
        public string InputToken { get; }
        public string OutputToken { get; }

        public virtual bool Execute(AutoCompleteEventArgs e)
        {
            if (!e.IsCharacter)
            {
                return false;
            }

            using (var pane = e.CodePane)
            {
                var selection = pane.Selection;
                if (selection.StartColumn < 2) { return false; }
                
                if (!e.IsCommitted && e.Character.ToString() == InputToken)
                {
                    using (var module = pane.CodeModule)
                    {
                        var newCode = e.OldCode.Insert(selection.StartColumn - 1, OutputToken);
                        module.ReplaceLine(selection.StartLine, newCode);
                        pane.Selection = selection;
                        e.NewCode = newCode;
                        return true;
                    }
                }
                return false;
            }
        }
    }
}
