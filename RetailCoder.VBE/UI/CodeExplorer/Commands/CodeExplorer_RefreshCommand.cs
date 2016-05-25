using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_RefreshCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorer_RefreshCommand(RubberduckParserState state)
        {
            _state = state;
        }

        public override void Execute(object parameter)
        {
            _state.OnParseRequested(this);
        }
    }
}
