using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Rubberduck.Parsing.VBA.Extensions;

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

        public bool IsCapturing { get; set; }

        protected virtual bool FindInputTokenAtBeginningOfCurrentLine => false;
        protected virtual bool SkipPreCompilerDirective => true;

        protected readonly IConfigProvider<IndenterSettings> IndenterSettings;

        protected virtual bool ExecuteOnCommittedInputOnly => true;
        protected virtual bool MatchInputTokenAtEndOfLineOnly => false;

        protected virtual bool IndentBody => true;

        public override bool Execute(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            var ignoreTab = e.Character == '\t' && !settings.CompleteBlockOnTab;
            var ignoreEnter = e.Character == '\r' && !settings.CompleteBlockOnEnter;
            if (IsInlineCharCompletion || e.IsDelete || ignoreTab || ignoreEnter)
            {
                return false;
            }

            var module = e.CodeModule;
            using (var pane = module.CodePane)
            {
                var selection = pane.Selection;
                var originalCode = module.GetLines(selection);
                var code = originalCode.Trim().StripStringLiterals();
                var hasComment = code.HasComment(out int commentStart);

                var isDeclareStatement = Regex.IsMatch(code, $"\\b{Tokens.Declare}\\b", RegexOptions.IgnoreCase);
                var isExitStatement = Regex.IsMatch(code, $"\\b{Tokens.Exit}\\b", RegexOptions.IgnoreCase);
                var isNamedArg = Regex.IsMatch(code, $"\\b{InputToken}\\:\\=", RegexOptions.IgnoreCase);

                if ((SkipPreCompilerDirective && code.StartsWith("#"))
                    || isDeclareStatement || isExitStatement || isNamedArg)
                {
                    return false;
                }

                if (IsMatch(code) && !IsBlockCompleted(module, selection))
                {
                    var indent = originalCode.TakeWhile(c => char.IsWhiteSpace(c)).Count();
                    var newCode = OutputToken.PadLeft(OutputToken.Length + indent, ' ');

                    var stdIndent = IndentBody
                        ? IndenterSettings.Create().IndentSpaces
                        : 0;

                    module.InsertLines(selection.NextLine.StartLine, "\n" + newCode);

                    module.ReplaceLine(selection.NextLine.StartLine, new string(' ', indent + stdIndent));
                    pane.Selection = new Selection(selection.NextLine.StartLine, indent + stdIndent + 1);

                    e.Handled = true;
                    return true;
                }
                return false;
            }
        }

        public override bool IsMatch(string code)
        {
            code = code.Trim().StripStringLiterals();
            var pattern = SkipPreCompilerDirective
                ? $"\\b{InputToken}\\b"
                : $"{InputToken}\\b"; // word boundary marker (\b) would prevent matching the # character

            bool regexOk;
            if (MatchInputTokenAtEndOfLineOnly)
            {
                regexOk = !code.StartsWith(Tokens.Else, System.StringComparison.OrdinalIgnoreCase) && 
                           code.EndsWith(InputToken, System.StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                regexOk = Regex.IsMatch(code, pattern, RegexOptions.IgnoreCase);
            }

            var hasComment = code.HasComment(out int commentIndex);
            return regexOk && (!hasComment || code.IndexOf(InputToken) < commentIndex);
        }

        protected bool IsBlockCompleted(ICodeModule module, Selection selection)
        {
            string content;
            var proc = module.GetProcOfLine(selection.StartLine);
            if (proc == null)
            {
                content = module.GetLines(1, module.CountOfDeclarationLines).StripStringLiterals();
            }
            else
            {
                var procKind = module.GetProcKindOfLine(selection.StartLine);
                var startLine = module.GetProcStartLine(proc, procKind);
                var lineCount = module.GetProcCountLines(proc, procKind);
                content = module.GetLines(startLine, lineCount);
            }

            var options = RegexOptions.IgnoreCase;
            var inputPattern = $"(?<!{OutputToken.Replace(InputToken, string.Empty)})\\b{InputToken}\\b";
            var inputMatches = Regex.Matches(content, inputPattern, options).Count;
            var outputMatches = Regex.Matches(content, $"\\b{OutputToken}\\b", options).Count;

            return inputMatches > 0 && inputMatches == outputMatches;
        }
    }
}
