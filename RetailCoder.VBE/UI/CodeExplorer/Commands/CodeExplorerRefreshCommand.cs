using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorerRefreshCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public CodeExplorerRefreshCommand(RubberduckParserState state)
        {
            _state = state;
        }

        public override void Execute(object parameter)
        {
            _state.OnParseRequested(this);
        }
    }
}