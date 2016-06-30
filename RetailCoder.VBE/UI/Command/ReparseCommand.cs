using System.Runtime.InteropServices;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.Command.MenuItems;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class ReparseCommandMenuItem : CommandMenuItemBase
    {
        public ReparseCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key
        {
            get { return "HotkeyDescription_ParseAll"; }
        }
    }

    [ComVisible(false)]
    public class ReparseCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public ReparseCommand(RubberduckParserState state)
        {
            _state = state;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.ParseAll; }
        }

        public override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Pending
                   || _state.Status == ParserState.Ready
                   || _state.Status == ParserState.Error
                   || _state.Status == ParserState.ResolverError;
        }

        public override void ExecuteImpl(object parameter)
        {
            _state.OnParseRequested(this);
        }
    }
}
