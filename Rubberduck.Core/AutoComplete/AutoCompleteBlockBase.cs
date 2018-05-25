using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBlockBase : AutoCompleteBase
    {
        /// <param name="inputToken">The token that starts the block, i.e. what to detect.</param>
        /// <param name="outputToken">The token that closes the block, i.e. what to insert.</param>
        protected AutoCompleteBlockBase(string inputToken, string outputToken)
            :base(inputToken, outputToken) { }

        public override bool Execute(AutoCompleteEventArgs e)
        {
            var selection = e.CodePane.Selection;

            if (e.IsCommitted && e.OldCode.Trim().StartsWith($"{InputToken} "))
            {
                var indent = e.OldCode.IndexOf(InputToken + 1); // borked
                using (var module = e.CodePane.CodeModule)
                {
                    var code = OutputToken.PadLeft($"{indent}{OutputToken}".Length, ' ');
                    module.InsertLines(selection.StartLine + 1, code);
                    e.CodePane.Selection = selection; // todo auto-indent?
                    e.NewCode = e.OldCode; // code at selection didn't change
                    return true;
                }
            }
            return false;
        }
    }
}
