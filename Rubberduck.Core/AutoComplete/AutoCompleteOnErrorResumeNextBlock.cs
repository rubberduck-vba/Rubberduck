using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteOnErrorResumeNextBlock : AutoCompleteBlockBase
    {
        public AutoCompleteOnErrorResumeNextBlock()
            : base($"{Tokens.On} {Tokens.Error} {Tokens.Resume} {Tokens.Next}", $"{Tokens.On} {Tokens.Error} {Tokens.GoTo} 0", false) { }
    }
}
