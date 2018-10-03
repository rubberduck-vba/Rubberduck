using Rubberduck.Parsing.Grammar;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteFunctionBlock : AutoCompleteBlockBase
    {
        public AutoCompleteFunctionBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Function}", $"{Tokens.End} {Tokens.Function}") { }

        public override bool Execute(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            var result = base.Execute(e, settings);
            if (result)
            {
                var module = e.Module;
                using (var pane = module.CodePane)
                {
                    var selection = pane.Selection;
                    var original = module.GetLines(selection);
                    var hasAsToken = Regex.IsMatch(original, $"\\)\\s+{Tokens.As}", RegexOptions.IgnoreCase) ||
                                     Regex.IsMatch(original, $"{Tokens.Function}\\s+\\(.*\\)\\s+{Tokens.As} ", RegexOptions.IgnoreCase);
                    var hasAsType = Regex.IsMatch(original, $"{Tokens.Function}\\s+\\w+\\(.*\\)\\s+{Tokens.As}\\s+\\w+", RegexOptions.IgnoreCase);
                    var asTypeClause =  hasAsToken && hasAsType
                        ? string.Empty 
                        : hasAsToken
                            ? $" {Tokens.Variant}"
                            : $" {Tokens.As} {Tokens.Variant}";

                    var code = original + asTypeClause;
                    module.ReplaceLine(selection.StartLine, code);
                    var newCode = module.GetLines(selection);
                    if (code == newCode)
                    {
                        pane.Selection = new Selection(selection.StartLine, code.Length - Tokens.Variant.Length + 1,
                                                       selection.StartLine, code.Length + 1);
                    }
                }
            }

            return result;
        }
    }
}
