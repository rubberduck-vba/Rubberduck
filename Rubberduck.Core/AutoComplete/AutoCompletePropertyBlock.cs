using Rubberduck.Parsing.Grammar;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete
{
    public class AutoCompletePropertyBlock : AutoCompleteBlockBase
    {
        public AutoCompletePropertyBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Property}", $"{Tokens.End} {Tokens.Property}") { }

        public override bool Execute(AutoCompleteEventArgs e, AutoCompleteSettings settings)
        {
            var result = base.Execute(e, settings);
            var module = e.CodeModule;
            using (var pane = module.CodePane)
            {
                var original = module.GetLines(e.CurrentSelection);
                var hasAsToken = Regex.IsMatch(original, $@"{Tokens.Property} {Tokens.Get}\s+\(.*\)\s+{Tokens.As}\s?", RegexOptions.IgnoreCase);
                var hasAsType = Regex.IsMatch(original, $@"{Tokens.Property} {Tokens.Get}\s+\w+\(.*\)\s+{Tokens.As}\s+(?<Identifier>\w+)", RegexOptions.IgnoreCase);
                var asTypeClause = hasAsToken && hasAsType
                    ? string.Empty
                    : hasAsToken
                        ? $" {Tokens.Variant}"
                        : $" {Tokens.As} {Tokens.Variant}";


                if (result && Regex.IsMatch(original, $"{Tokens.Property} {Tokens.Get}"))
                {
                    var code = original + asTypeClause;
                    module.ReplaceLine(e.CurrentSelection.StartLine, code);
                    var newCode = module.GetLines(e.CurrentSelection);
                    if (code == newCode)
                    {
                        pane.Selection = new Selection(e.CurrentSelection.StartLine, code.Length - Tokens.Variant.Length + 1,
                                                        e.CurrentSelection.StartLine, code.Length + 1);
                    }
                }
            }

            return result;
        }

    }
}
