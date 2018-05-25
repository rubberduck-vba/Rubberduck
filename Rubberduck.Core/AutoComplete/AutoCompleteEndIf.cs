using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteEndIf : AutoCompleteBase
    {
        public override string InputToken => "If ";
        public override string OutputToken => "End If";

        public override string Execute(AutoCompleteEventArgs e)
        {
            var selection = e.CodePane.Selection;

            if (e.IsCommitted && e.Code.Trim().StartsWith(InputToken))
            {
                var indent = e.Code.IndexOf(InputToken + 1); // borked
                using (var module = e.CodePane.CodeModule)
                {
                    var code = OutputToken.PadLeft(indent + OutputToken.Length, ' ');
                    module.InsertLines(selection.StartLine + 1, code);
                    e.CodePane.Selection = selection; // todo auto-indent?
                    return code;
                }
            }
            return null;
        }
    }
}
