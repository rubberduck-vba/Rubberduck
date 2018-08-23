using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.If}", $"{Tokens.End} {Tokens.If}") { }

        // matching "If" would trigger erroneous block completion on inline if..then..else syntax.
        protected override bool MatchInputTokenAtEndOfLineOnly => true;
    }
}
