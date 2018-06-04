using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.Events;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete
{
    public abstract class AutoCompleteBlockBase : AutoCompleteBase
    {
        /// <param name="api">Used for ensuring compilable resulting code.</param>
        /// <param name="indenterSettings">Used for auto-indenting blocks as per indenter settings.</param>
        /// <param name="inputToken">The token that starts the block, i.e. what to detect.</param>
        /// <param name="outputToken">The token that closes the block, i.e. what to insert.</param>
        /// <param name="committedOnly">Indicates whether line of code was committed, i.e. selection is on the line underneath the code string.</param>
        protected AutoCompleteBlockBase(IVBETypeLibsAPI api, IIndenterSettings indenterSettings, string inputToken, string outputToken, bool committedOnly = true)
            :base(inputToken, outputToken)
        {
            _api = api;
            _indenterSettings = indenterSettings;
            _committedOnly = committedOnly;
        }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;

        private readonly IVBETypeLibsAPI _api;
        private readonly IIndenterSettings _indenterSettings;
        private readonly bool _committedOnly;

        protected virtual bool ExecuteOnCommittedInputOnly => true;
        protected virtual bool MatchInputTokenAtEndOfLineOnly => false;

        private string _pattern => MatchInputTokenAtEndOfLineOnly
                                    ? $"\\b{InputToken}\\r\\n$"
                                    : $"\\b{InputToken}\\b";

        private bool _executing;
        public override bool Execute(AutoCompleteEventArgs e)
        {
            if (_executing)
            {
                return false;
            }

            var selection = e.CodePane.Selection;
            var stdIndent = _indenterSettings.IndentSpaces;

            if ((!_committedOnly || e.IsCommitted) && Regex.IsMatch(e.OldCode.Trim(), _pattern))
            {
                var indent = e.OldCode.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                using (var module = e.CodePane.CodeModule)
                {
                    _executing = true;
                    var code = OutputToken.PadLeft(OutputToken.Length + indent, ' ');
                    if (module.GetLines(selection.NextLine) == code)
                    {
                        return false;
                    }

                    module.InsertLines(selection.StartLine + 1, code);

                    module.ReplaceLine(selection.StartLine, new string(' ', indent + stdIndent));
                    e.CodePane.Selection = new VBEditor.Selection(selection.StartLine, indent + stdIndent + 1);
                    e.NewCode = e.OldCode;
                    _executing = false;
                    return true;
                }
            }
            return false;
        }
    }
}
