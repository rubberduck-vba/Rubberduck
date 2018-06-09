using Rubberduck.Parsing.VBA;
using Rubberduck.SettingsProvider;
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
        protected AutoCompleteBlockBase(IConfigProvider<IndenterSettings> indenterSettings, string inputToken, string outputToken)
            :base(inputToken, outputToken)
        {
            IndenterSettings = indenterSettings;
        }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;
        protected virtual bool SkipPreCompilerDirective => true;

        protected readonly IConfigProvider<IndenterSettings> IndenterSettings;

        protected virtual bool ExecuteOnCommittedInputOnly => true;
        protected virtual bool MatchInputTokenAtEndOfLineOnly => false;

        protected virtual bool IndentBody => true;

        public override bool Execute(AutoCompleteEventArgs e)
        {
            if (e.Character.ToString() != "\r")
            {
                // handle ENTER or TAB key
                return false;
            }

            using (var pane = e.CodePane)
            using (var module = pane.CodeModule)
            {
                var selection = pane.Selection;
                var code = module.GetLines(selection);

                if (SkipPreCompilerDirective && code.Trim().StartsWith("#"))
                {
                    return false;
                }

                var pattern = SkipPreCompilerDirective
                                ? $"\\b{InputToken}\\b"
                                : $"{InputToken}\\b"; // word boundary marker (\b) would prevent matching the # character

                var isMatch = MatchInputTokenAtEndOfLineOnly
                                ? code.EndsWith(InputToken)
                                : Regex.IsMatch(code.Trim(), pattern);

                if (!code.HasComment(out _) && isMatch)
                {
                    var indent = code.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                    var newCode = OutputToken.PadLeft(OutputToken.Length + indent, ' ');
                    if (module.GetLines(selection.NextLine) == newCode)
                    {
                        return false;
                    }

                    var stdIndent = IndentBody ? IndenterSettings.Create().IndentSpaces : 0;

                    module.InsertLines(selection.NextLine.StartLine+1, newCode);

                    module.ReplaceLine(selection.NextLine.StartLine, new string(' ', indent + stdIndent));
                    e.CodePane.Selection = new VBEditor.Selection(selection.NextLine.StartLine, indent + stdIndent + 1);

                    e.NewCode = newCode;
                    return true;
                }
                return false;
            }
        }
    }
}
