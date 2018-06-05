using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompletePrecompilerIfBlock : AutoCompleteBlockBase
    {
        public AutoCompletePrecompilerIfBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"#{Tokens.If}", $"#{Tokens.End} {Tokens.If}") { }

        protected override bool SkipPreCompilerDirective => false;
    }
}
