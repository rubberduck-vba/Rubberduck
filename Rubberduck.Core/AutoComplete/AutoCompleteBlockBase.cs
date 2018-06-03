using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Events;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBlockBase : AutoCompleteBase
    {
        /// <param name="indenterSettings">Used for auto-indenting blocks as per indenter settings.</param>
        /// <param name="inputToken">The token that starts the block, i.e. what to detect.</param>
        /// <param name="outputToken">The token that closes the block, i.e. what to insert.</param>
        /// <param name="committedOnly">Indicates whether line of code was committed, i.e. selection is on the line underneath the code string.</param>
        protected AutoCompleteBlockBase(IIndenterSettings indenterSettings, string inputToken, string outputToken, bool committedOnly = true)
            :base(inputToken, outputToken)
        {
            _indenterSettings = indenterSettings;
            _committedOnly = committedOnly;
        }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;

        private readonly IIndenterSettings _indenterSettings;
        private readonly bool _committedOnly;

        public override bool Execute(AutoCompleteEventArgs e)
        {
            var selection = e.CodePane.Selection;
            var stdIndent = _indenterSettings.IndentSpaces;

            if ((!_committedOnly || e.IsCommitted) && Regex.IsMatch(e.OldCode.Trim(), $"\\b{InputToken}\\b"))
            {
                var indent = e.OldCode.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                using (var module = e.CodePane.CodeModule)
                {
                    var code = OutputToken.PadLeft(OutputToken.Length + indent, ' ');
                    module.InsertLines(selection.StartLine + 1, code);
                    module.ReplaceLine(selection.StartLine, new string(' ', indent + stdIndent));
                    e.CodePane.Selection = new VBEditor.Selection(selection.StartLine, indent + stdIndent + 1);
                    e.NewCode = e.OldCode;
                    return true;
                }
            }
            return false;
        }
    }
}
