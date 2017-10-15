using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.CodeExplorer.Commands;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    [CodeExplorerCommand]
    public class ReparseCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public ReparseCommand(RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.ParseAll; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Pending
                   || _state.Status == ParserState.Ready
                   || _state.Status == ParserState.Error
                   || _state.Status == ParserState.ResolverError;
        }

        protected override void ExecuteImpl(object parameter)
        {
            _state.OnParseRequested(this);
        }
    }
}
