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
        protected AutoCompleteBlockBase(IVBETypeLibsAPI api, IIndenterSettings indenterSettings, string inputToken, string outputToken)
            :base(inputToken, outputToken)
        {
            _api = api;
            _indenterSettings = indenterSettings;
        }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;

        private readonly IVBETypeLibsAPI _api;
        private readonly IIndenterSettings _indenterSettings;

        protected virtual bool ExecuteOnCommittedInputOnly => true;
        protected virtual bool MatchInputTokenAtEndOfLineOnly => false;

        private bool _executing;
        public override bool Execute(AutoCompleteEventArgs e)
        {
            if (_executing)
            {
                return false;
            }

            var selection = e.CodePane.Selection;
            var stdIndent = _indenterSettings.IndentSpaces;

            var isMatch = MatchInputTokenAtEndOfLineOnly 
                            ? e.OldCode.EndsWith(InputToken)
                            : Regex.IsMatch(e.OldCode.Trim(), $"\\b{InputToken}\\b");

            if (isMatch && (!ExecuteOnCommittedInputOnly || e.IsCommitted))
            {
                var indent = e.OldCode.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                using (var module = e.CodePane.CodeModule)
                {
                    _executing = true;
                    var code = OutputToken.PadLeft(OutputToken.Length + indent, ' ');
                    if (module.GetLines(selection.NextLine) == code)
                    {
                        _executing = false;
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
