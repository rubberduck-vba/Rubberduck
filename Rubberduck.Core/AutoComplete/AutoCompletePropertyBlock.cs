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
                if (result && Regex.IsMatch(original, $"{Tokens.Property} {Tokens.Get}"))
                {
                    var asTypeClause = $" {Tokens.As} {Tokens.Variant}";
                    var code = original + (Regex.IsMatch(original, $"\\) {Tokens.As} ") ? string.Empty : asTypeClause);
                    module.ReplaceLine(e.CurrentSelection.StartLine, code);
                    pane.Selection = new Selection(e.CurrentSelection.StartLine, code.Length - Tokens.Variant.Length + 1,
                                                    e.CurrentSelection.StartLine, code.Length + 1);
                }
            }

            return result;
        }

    }
}
