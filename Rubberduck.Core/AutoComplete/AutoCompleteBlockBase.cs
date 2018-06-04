using Rubberduck.Parsing.VBA;
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
        protected AutoCompleteBlockBase(IIndenterSettings indenterSettings, string inputToken, string outputToken)
            :base(inputToken, outputToken)
        {
            IndenterSettings = indenterSettings;
        }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;
        protected virtual bool SkipPreCompilerDirective => true;

        protected readonly IIndenterSettings IndenterSettings;

        protected virtual bool ExecuteOnCommittedInputOnly => true;
        protected virtual bool MatchInputTokenAtEndOfLineOnly => false;

        protected virtual bool IndentBody => true;

        private bool _executing;
        public override bool Execute(AutoCompleteEventArgs e)
        {
            if (_executing || (SkipPreCompilerDirective && e.OldCode.Trim().StartsWith("#")))
            {
                return false;
            }

            var selection = e.CodePane.Selection;
            var stdIndent = IndentBody ? IndenterSettings.IndentSpaces : 0;

            var pattern = SkipPreCompilerDirective
                            ? $"\\b{InputToken}\\b"
                            : $"{InputToken}\\b"; // word boundary marker (\b) would prevent matching the # character

            var isMatch = MatchInputTokenAtEndOfLineOnly 
                            ? e.OldCode.EndsWith(InputToken)
                            : Regex.IsMatch(e.OldCode.Trim(), pattern);

            if (!e.OldCode.HasComment(out _) && isMatch && (!ExecuteOnCommittedInputOnly || e.IsCommitted))
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
